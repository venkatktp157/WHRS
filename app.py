import io
import numpy as np
import openpyxl
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from statsmodels.regression.linear_model import OLS

# Page setup
st.set_page_config(
    page_title="Waster Heat Recovery System ROI Calculator",
    page_icon="⚓",
    layout="wide",
)

st.title("⚓ Waste Heat Recovery System- ROI Calculator")
st.markdown("""
Upload your WHRS Excel log file or CSV containing raw input columns. 
The application automatically processes all metrics, applies multi-type filters, and provides customizable time-series and correlation visualizations.
""")

# Standard fuel lookup tables from sheet formulas
FUEL_PRICE_MAP = {"LNG": 1345.0, "LSMGO": 1280.0, "VLSFO": 820.0, "HSFO": 635.0}


def compute_whrs_metrics(df_in):
  """Replicates all Excel formulas and conditional logic while preventing column duplication."""
  df = df_in.copy()

  # Clean column headers
  df.columns = [str(c).strip() for c in df.columns]

  # Required column name mappings
  net_total_col = next(
      (c for c in df.columns if "NET KW TOTAL" in c and "RECOVERED" not in c),
      "WHRS NET KW TOTAL",
  )
  gross_total_col = next(
      (
          c
          for c in df.columns
          if "KW TOTAL" in c and "NET" not in c and "RECOVERED" not in c
      ),
      "WHRS KW TOTAL",
  )
  fuel_col = next(
      (c for c in df.columns if "ME FUEL IN USE" in c), "ME FUEL IN USE"
  )
  rh_col = next((c for c in df.columns if "WHRS RH" in c), "WHRS RH")

  # Clean fuel column string safely
  if fuel_col in df.columns:
    fuel_series = (
        df[fuel_col]
        .astype(str)
        .str.strip()
        .str.upper()
        .replace({"NAN": "", "NONE": ""})
    )
  else:
    fuel_series = pd.Series([""] * len(df))

  # 1. Gen SFOC (g/kWh) Formula (Blank if fuel is blank)
  df["Gen SFOC (g/kWh)"] = np.where(
      fuel_series == "",
      np.nan,
      np.where(fuel_series == "LNG", 192.1, 182.4),
  )

  # 2. Fuel Price ($/mt) Formula (Blank if fuel is blank)
  df["Fuel Price ($/mt)"] = fuel_series.map(FUEL_PRICE_MAP)

  # Convert numeric columns to float safely (Keep NaN as NaN)
  for col in [net_total_col, gross_total_col, rh_col]:
    if col in df.columns:
      df[col] = pd.to_numeric(df[col], errors="coerce")
    else:
      df[col] = np.nan

  # Calculate difference metrics
  net_diff = df[net_total_col].diff()
  net_diff.iloc[0] = df[net_total_col].iloc[0]
  calculated_net_recovered = net_diff.clip(lower=0)

  gross_diff = df[gross_total_col].diff()
  gross_diff.iloc[0] = df[gross_total_col].iloc[0]
  calculated_gross_recovered = gross_diff.clip(lower=0)

  # Prevent duplicate columns by dropping pre-existing ones
  cols_to_clean = [
      "WHRS NET KW RECOVERED PER DAY",
      "WHRS  NET KW RECOVERED PER DAY",
      "WHRS KW RECOVERED PER DAY",
  ]
  df = df.loc[:, ~df.columns.isin(cols_to_clean)]

  # Assign calculated metrics
  df["WHRS NET KW RECOVERED PER DAY"] = calculated_net_recovered
  df["WHRS KW RECOVERED PER DAY"] = calculated_gross_recovered

  # 5. WHRS KWh (Average power generation per running hour)
  kwh_condition = (
      (df["WHRS KW RECOVERED PER DAY"].fillna(0) != 0)
      & (df[rh_col].fillna(0) != 0)
  )

  df["WHRS KWh"] = np.where(
      kwh_condition,
      (df["WHRS KW RECOVERED PER DAY"] / df[rh_col]).round(1),
      0.0,
  )

  # 6. Daily ROI (USD)
  missing_condition = (
      df["WHRS NET KW RECOVERED PER DAY"].isna()
      | df["Gen SFOC (g/kWh)"].isna()
      | df["Fuel Price ($/mt)"].isna()
  )

  df["Daily ROI (USD)"] = np.where(
      missing_condition,
      np.nan,  # Returns NaN for proper numerical aggregation
      (
          df["WHRS NET KW RECOVERED PER DAY"]
          * (df["Gen SFOC (g/kWh)"] / 1000000.0)
          * df["Fuel Price ($/mt)"]
      ),
  )

  # 7. Total ROI (USD) - Cumulative
  daily_roi_numeric = pd.to_numeric(df["Daily ROI (USD)"], errors="coerce")
  cum_sum = daily_roi_numeric.cumsum()

  # Returns NaN on blank daily ROI rows
  df["Total ROI (USD)"] = np.where(daily_roi_numeric.isna(), np.nan, cum_sum)

  return df


# -----------------------------------------------------------------------------
# FILE UPLOAD & PROCESSING UI
# -----------------------------------------------------------------------------
uploaded_file = st.file_uploader(
    "Upload Excel (.xlsx, .xls) or CSV dataset", type=["xlsx", "xls", "csv"]
)

if uploaded_file is not None:
  try:
    # Load dataset based on file type
    if uploaded_file.name.endswith(".csv"):
      raw_df = pd.read_csv(uploaded_file)
    else:
      xls = pd.ExcelFile(uploaded_file)
      sheet_name = xls.sheet_names[0]

      # Auto-locate header row
      preview_df = pd.read_excel(
          uploaded_file, sheet_name=sheet_name, nrows=5, header=None
      )
      header_row = 0
      for idx, row in preview_df.iterrows():
        row_str = " ".join([str(val) for val in row.values])
        if any(
            k in row_str
            for k in [
                "ME FUEL IN USE",
                "DATE TIME",
                "VESSEL STATUS",
                "DATE",
            ]
        ):
          header_row = idx
          break

      raw_df = pd.read_excel(
          uploaded_file, sheet_name=sheet_name, header=header_row
      )

    # Run automated calculations
    computed_df = compute_whrs_metrics(raw_df)

    # -----------------------------------------------------------------------------
    # DYNAMIC PARAMETER & DATE FILTERING
    # -----------------------------------------------------------------------------
    st.sidebar.header("🗓️ Data & Parameter Filters")

    # Automatically detect the primary datetime column for plot X-axes
    datetime_cols = [
        c
        for c in computed_df.columns
        if any(k in c.upper() for k in ["DATE", "TIME", "TIMESTAMP"])
    ]
    selected_dt_col = datetime_cols[0] if datetime_cols else None

    # Sidebar filter selection
    selected_filter_col = st.sidebar.selectbox(
        "Select Parameter / Column to Filter",
        options=computed_df.columns,
        index=(
            computed_df.columns.get_loc(selected_dt_col)
            if selected_dt_col
            else 0
        ),
    )

    filtered_df = computed_df.copy()

    if selected_filter_col:
      col_data = filtered_df[selected_filter_col]

      # Check if column is Datetime
      is_datetime = False
      parsed_dates = pd.to_datetime(col_data, errors="coerce")
      if parsed_dates.notna().sum() > 0 and any(
          k in selected_filter_col.upper()
          for k in ["DATE", "TIME", "TIMESTAMP"]
      ):
        is_datetime = True

      # --- CASE A: DATETIME FILTERING ---
      if is_datetime:
        selected_dt_col = selected_filter_col
        valid_dates = parsed_dates.dropna()
        filtered_df["_parsed_datetime"] = parsed_dates
        min_date = valid_dates.min().date()
        max_date = valid_dates.max().date()

        if min_date != max_date:
          date_range = st.sidebar.date_input(
              f"Filter Date Range ({selected_filter_col})",
              value=(min_date, max_date),
              min_value=min_date,
              max_value=max_date,
          )

          if isinstance(date_range, tuple) and len(date_range) == 2:
            start_date, end_date = date_range
            mask = (filtered_df["_parsed_datetime"].dt.date >= start_date) & (
                filtered_df["_parsed_datetime"].dt.date <= end_date
            )
            filtered_df = filtered_df[mask]

        filtered_df = filtered_df.drop(
            columns=["_parsed_datetime"], errors="ignore"
        )

      # --- CASE B: NUMERIC FILTERING (Slider) ---
      elif pd.api.types.is_numeric_dtype(col_data):
        min_val = float(col_data.min())
        max_val = float(col_data.max())

        if min_val != max_val:
          selected_range = st.sidebar.slider(
              f"Filter Range for {selected_filter_col}",
              min_value=min_val,
              max_value=max_val,
              value=(min_val, max_val),
          )
          filtered_df = filtered_df[
              (filtered_df[selected_filter_col] >= selected_range[0])
              & (filtered_df[selected_filter_col] <= selected_range[1])
          ]

      # --- CASE C: CATEGORICAL / TEXT FILTERING (Multiselect) ---
      else:
        unique_values = col_data.dropna().astype(str).unique().tolist()
        selected_options = st.sidebar.multiselect(
            f"Select values for {selected_filter_col}",
            options=unique_values,
            default=unique_values,
        )
        if selected_options:
          filtered_df = filtered_df[
              filtered_df[selected_filter_col].astype(str).isin(selected_options)
          ]

    st.success(f"File **{uploaded_file.name}** processed successfully!")

    # -----------------------------------------------------------------------------
    # OVERVIEW METRICS (SAFE NUMERIC COMPUTATION)
    # -----------------------------------------------------------------------------
    total_rows = len(filtered_df)
    total_net_power = (
        pd.to_numeric(
            filtered_df.get("WHRS NET KW RECOVERED PER DAY", 0), errors="coerce"
        ).sum()
        if "WHRS NET KW RECOVERED PER DAY" in filtered_df
        else 0.0
    )
    total_roi = (
        pd.to_numeric(
            filtered_df.get("Daily ROI (USD)", 0), errors="coerce"
        ).sum()
        if "Daily ROI (USD)" in filtered_df
        else 0.0
    )

    col1, col2, col3 = st.columns(3)
    col1.metric("Filtered Records", f"{total_rows}")
    col2.metric("Total Net Energy Recovered", f"{total_net_power:,.1f} kWh")
    col3.metric("Total Financial ROI Savings", f"${total_roi:,.2f} USD")

    # -----------------------------------------------------------------------------
    # VISUALIZATIONS SECTION
    # -----------------------------------------------------------------------------
    st.divider()
    st.subheader("📈 Visualization & Trend Analysis")

    viz_tab1, viz_tab2 = st.tabs([
        "⏱️ Individual Time Series Trends",
        "🔀 Flexible X vs. Y Scatter / Correlation",
    ])

    # Helper function to find best column matches
    def find_col_match(df, keywords):
      for kw in keywords:
        for c in df.columns:
          if kw.upper() in c.upper():
            return c
      return None

    mapped_vars = {
        "Gen SFOC (g/kWh)": find_col_match(filtered_df, ["Gen SFOC", "SFOC"]),
        "WHRS Net KW Recovered / Day": find_col_match(
            filtered_df, ["NET KW RECOVERED", "WHRS NET KW RECOVERED PER DAY"]
        ),
        "WHRS Gross KW Recovered / Day": find_col_match(
            filtered_df, ["WHRS KW RECOVERED PER DAY", "GROSS KW RECOVERED"]
        ),
        "WHRS Run Hours (WHRS RH)": find_col_match(
            filtered_df, ["WHRS RH", "RUN HRS", "RUNNING HOURS"]
        ),
        "WHRS KWh (Avg Power / RH)": find_col_match(filtered_df, ["WHRS KWH"]),
        "Exhaust Gas Bypass (%)": find_col_match(
            filtered_df, ["BYPASS", "EXHAUST BYPASS", "EGB"]
        ),
        "Engine Load (%)": find_col_match(
            filtered_df, ["LOAD", "ME LOAD", "ENGINE LOAD"]
        ),
        "JW Flow Rate (m3/hr)": find_col_match(
            filtered_df, ["JW FLOW", "JACKET WATER", "JW RATE"]
        ),
        "Daily & Cumulative Total ROI ($)": "ROI_COMBO",
    }

    numeric_cols = (
        filtered_df.select_dtypes(include=[np.number]).columns.tolist()
    )

    # -----------------------------------------------------------------------------
    # TAB 1: INDIVIDUAL TIME SERIES TRENDS
    # -----------------------------------------------------------------------------
    with viz_tab1:
      st.markdown("#### Select Y-Axis Metric to Plot Over Time")

      metric_options = [
          k
          for k, v in mapped_vars.items()
          if v is not None or k == "Daily & Cumulative Total ROI ($)"
      ]

      for nc in numeric_cols:
        if nc not in mapped_vars.values():
          metric_options.append(nc)

      selected_metric = st.selectbox(
          "Select Y-Axis Metric for Time Series:",
          options=metric_options,
          index=0,
      )

      x_data = (
          filtered_df[selected_dt_col]
          if (selected_dt_col and selected_dt_col in filtered_df.columns)
          else filtered_df.index
      )

      # Case 1: Daily ROI & Total ROI Combined Plot
      if selected_metric == "Daily & Cumulative Total ROI ($)":
        fig_ts = go.Figure()

        if "Daily ROI (USD)" in filtered_df.columns:
          fig_ts.add_trace(
              go.Bar(
                  x=x_data,
                  y=pd.to_numeric(
                      filtered_df["Daily ROI (USD)"], errors="coerce"
                  ),
                  name="Daily ROI ($)",
                  marker_color="#2b5c8f",
                  opacity=0.75,
              )
          )

        if "Total ROI (USD)" in filtered_df.columns:
          fig_ts.add_trace(
              go.Scatter(
                  x=x_data,
                  y=pd.to_numeric(
                      filtered_df["Total ROI (USD)"], errors="coerce"
                  ),
                  name="Cumulative Total ROI ($)",
                  mode="lines+markers",
                  line=dict(color="#2ca02c", width=3),
                  yaxis="y2",
              )
          )

        fig_ts.update_layout(
            title="Daily ROI ($) and Cumulative Total ROI ($) Over Time",
            xaxis_title=(
                selected_dt_col if selected_dt_col else "Log Entry Index"
            ),
            yaxis=dict(title="Daily ROI (USD)"),
            yaxis2=dict(
                title="Cumulative Total ROI (USD)",
                overlaying="y",
                side="right",
            ),
            hovermode="x unified",
            template="plotly_white",
            height=500,
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="right",
                x=1,
            ),
        )

      # Case 2: Standard Single Metric Time Series
      else:
        actual_col = mapped_vars.get(selected_metric, selected_metric)

        if actual_col and actual_col in filtered_df.columns:
          fig_ts = px.line(
              filtered_df,
              x=x_data,
              y=actual_col,
              markers=True,
              title=f"{selected_metric} Over Time",
              labels={actual_col: selected_metric, "x": "Datetime / Entry"},
              template="plotly_white",
          )

          fig_ts.update_traces(line=dict(width=2, color="#1f77b4"))
          fig_ts.update_layout(hovermode="x unified", height=500)
        else:
          st.warning(
              f"Column **{selected_metric}** not found in the uploaded file."
          )
          fig_ts = None

      if fig_ts:
        st.plotly_chart(fig_ts, use_container_width=True)

    # -----------------------------------------------------------------------------
    # TAB 2: FLEXIBLE X VS. Y SCATTER / CORRELATION PLOT
    # -----------------------------------------------------------------------------
    with viz_tab2:
      st.markdown("#### Flexible Correlation Plot (Select X and Y Axis)")

      c_x, c_y, c_color = st.columns(3)

      load_default_col = mapped_vars["Engine Load (%)"] or (
          numeric_cols[0] if numeric_cols else None
      )
      net_default_col = mapped_vars["WHRS Net KW Recovered / Day"] or (
          numeric_cols[1] if len(numeric_cols) > 1 else numeric_cols[0]
      )

      default_x_idx = (
          numeric_cols.index(load_default_col)
          if load_default_col in numeric_cols
          else 0
      )
      default_y_idx = (
          numeric_cols.index(net_default_col)
          if net_default_col in numeric_cols
          else min(1, len(numeric_cols) - 1)
      )

      with c_x:
        selected_x = st.selectbox(
            "Select X-Axis Variable:",
            options=numeric_cols,
            index=default_x_idx,
            key="corr_x_select",
        )

      with c_y:
        selected_y = st.selectbox(
            "Select Y-Axis Variable:",
            options=numeric_cols,
            index=default_y_idx,
            key="corr_y_select",
        )

      with c_color:
        color_option = st.selectbox(
            "Color Data Points By (Optional):",
            options=["None"]
            + numeric_cols
            + [
                c
                for c in filtered_df.columns
                if filtered_df[c].dtype == "object"
            ],
            index=0,
            key="corr_color_select",
        )

      show_trendline = st.checkbox("Show OLS Trendline", value=False)

      if selected_x and selected_y:
        color_arg = color_option if color_option != "None" else None
        hover_cols = (
            [selected_dt_col]
            if (selected_dt_col and selected_dt_col in filtered_df.columns)
            else None
        )

        fig_scatter = px.scatter(
            filtered_df,
            x=selected_x,
            y=selected_y,
            color=color_arg,
            trendline="ols" if show_trendline else None,
            hover_data=hover_cols,
            title=f"Correlation: {selected_x} vs. {selected_y}",
            template="plotly_white",
        )

        fig_scatter.update_layout(height=520)
        st.plotly_chart(fig_scatter, use_container_width=True)

    # -----------------------------------------------------------------------------
    # DATA PREVIEW & EXPORT
    # -----------------------------------------------------------------------------
    st.divider()
    st.subheader("📊 Calculated Dataset Preview")
    st.dataframe(filtered_df, use_container_width=True)

    st.divider()
    st.subheader("📥 Export Processed Dataset")

    exp_col1, exp_col2 = st.columns(2)

    # Export as Excel
    buffer_excel = io.BytesIO()
    with pd.ExcelWriter(buffer_excel, engine="openpyxl") as writer:
      filtered_df.to_excel(
          writer, index=False, sheet_name="WHRS_Calculated_ROI"
      )
    buffer_excel.seek(0)

    exp_col1.download_button(
        label="📄 Download Clean Excel (.xlsx)",
        data=buffer_excel,
        file_name="WHRS_LOG_Calculated_ROI.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    )

    # Export as CSV
    buffer_csv = filtered_df.to_csv(index=False).encode("utf-8")
    exp_col2.download_button(
        label="📑 Download Clean CSV (.csv)",
        data=buffer_csv,
        file_name="WHRS_LOG_Calculated_ROI.csv",
        mime="text/csv",
    )

  except Exception as e:
    st.error(f"Error processing the uploaded file: {str(e)}")

else:
  st.info(
      "👆 Please upload your WHRS Excel log file or CSV to execute"
      " calculations and view trends."
  )