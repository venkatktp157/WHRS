import streamlit as st
import pandas as pd
import numpy as np
import io
import plotly.express as px
import plotly.graph_objects as go

# Page setup
st.set_page_config(
    page_title="WHRS Batch Calculator & ROI Processor",
    page_icon="⚓",
    layout="wide"
)

st.title("⚓ WHRS Log Batch File Processor & ROI Calculator")
st.markdown("""
Upload your WHRS Excel log file or CSV containing raw input columns. 
The application will automatically initialize all required columns, run calculations, allow date filtering, and render analytical visual trends.
""")

# Standard fuel lookup tables from sheet formulas
FUEL_PRICE_MAP = {
    "LNG": 1345.0,
    "LSMGO": 1280.0,
    "VLSFO": 820.0,
    "HSFO": 635.0
}

def compute_whrs_metrics(df_in):
    """Replicates all Excel formulas and conditional logic while preventing column duplication."""
    df = df_in.copy()
    
    # Clean column headers
    df.columns = [str(c).strip() for c in df.columns]

    # Required column name mappings (handling slight header variations)
    net_total_col = next((c for c in df.columns if "NET KW TOTAL" in c and "RECOVERED" not in c), "WHRS NET KW TOTAL")
    gross_total_col = next((c for c in df.columns if "KW TOTAL" in c and "NET" not in c and "RECOVERED" not in c), "WHRS KW TOTAL")
    fuel_col = next((c for c in df.columns if "ME FUEL IN USE" in c), "ME FUEL IN USE")
    rh_col = next((c for c in df.columns if "WHRS RH" in c), "WHRS RH")

    # Clean fuel column string
    fuel_series = df[fuel_col].astype(str).str.strip().str.upper() if fuel_col in df.columns else pd.Series(["VLSFO"]*len(df))

    # 1. Gen SFOC (g/kWh) Formula
    df["Gen SFOC (g/kWh)"] = np.where(fuel_series == "LNG", 192.1, 182.4)

    # 2. Fuel Price ($/mt) Formula
    df["Fuel Price ($/mt)"] = fuel_series.map(FUEL_PRICE_MAP).fillna(820.0)

    # Convert numeric columns to float safely
    for col in [net_total_col, gross_total_col, rh_col]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        else:
            df[col] = 0.0

    # Calculate difference metrics
    net_diff = df[net_total_col].diff()
    net_diff.iloc[0] = df[net_total_col].iloc[0]
    calculated_net_recovered = net_diff.clip(lower=0)

    gross_diff = df[gross_total_col].diff()
    gross_diff.iloc[0] = df[gross_total_col].iloc[0]
    calculated_gross_recovered = gross_diff.clip(lower=0)

    # Prevent duplicate columns by dropping pre-existing ones with matching names
    cols_to_clean = ["WHRS NET KW RECOVERED PER DAY", "WHRS  NET KW RECOVERED PER DAY", "WHRS KW RECOVERED PER DAY"]
    df = df.loc[:, ~df.columns.isin(cols_to_clean)]

    # Assign calculated metrics
    df["WHRS NET KW RECOVERED PER DAY"] = calculated_net_recovered
    df["WHRS KW RECOVERED PER DAY"] = calculated_gross_recovered

    # 5. WHRS KWh (Average power generation per running hour)
    kwh_condition = (
        (df["WHRS KW RECOVERED PER DAY"] != 0) & 
        (df["WHRS KW RECOVERED PER DAY"].notna()) & 
        (df[rh_col] != 0) & 
        (df[rh_col].notna())
    )

    df["WHRS KWh"] = np.where(
        kwh_condition,
        (df["WHRS KW RECOVERED PER DAY"] / df[rh_col]).round(1),
        0.0
    )

    # 6. Daily ROI (USD)
    df["Daily ROI (USD)"] = (
        df["WHRS NET KW RECOVERED PER DAY"] * 
        (df["Gen SFOC (g/kWh)"] / 1000000.0) * 
        df["Fuel Price ($/mt)"]
    )

    # 7. Total ROI (USD)
    df["Total ROI (USD)"] = df["Daily ROI (USD)"].cumsum()

    # Remove any duplicate columns
    df = df.loc[:, ~df.columns.duplicated()]

    return df


# -----------------------------------------------------------------------------
# FILE UPLOAD & PROCESSING UI
# -----------------------------------------------------------------------------
uploaded_file = st.file_uploader(
    "Upload Excel (.xlsx, .xls) or CSV dataset", 
    type=["xlsx", "xls", "csv"]
)

if uploaded_file is not None:
    try:
        # Load dataset based on file type
        if uploaded_file.name.endswith(".csv"):
            raw_df = pd.read_csv(uploaded_file)
        else:
            xls = pd.ExcelFile(uploaded_file)
            sheet_name = xls.sheet_names[0]
            
            # Read first few rows to locate column header row automatically
            preview_df = pd.read_excel(uploaded_file, sheet_name=sheet_name, nrows=5, header=None)
            header_row = 0
            for idx, row in preview_df.iterrows():
                row_str = " ".join([str(val) for val in row.values])
                if "ME FUEL IN USE" in row_str or "DATE TIME" in row_str or "VESSEL STATUS" in row_str or "DATE" in row_str:
                    header_row = idx
                    break
            
            raw_df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=header_row)

        # Run automated calculations
        computed_df = compute_whrs_metrics(raw_df)

        # -----------------------------------------------------------------------------
        # DATE / TIME FILTERING
        # -----------------------------------------------------------------------------
        st.sidebar.header("🗓️ Data Filters")
        
        # Try finding a datetime column automatically
        datetime_cols = [c for c in computed_df.columns if any(k in c.upper() for k in ["DATE", "TIME", "TIMESTAMP"])]
        
        selected_dt_col = st.sidebar.selectbox(
            "Select Date/Time Column",
            options=computed_df.columns,
            index=computed_df.columns.get_loc(datetime_cols[0]) if datetime_cols else 0
        )

        filtered_df = computed_df.copy()

        if selected_dt_col:
            # Parse to datetime
            parsed_dates = pd.to_datetime(filtered_df[selected_dt_col], errors='coerce')
            valid_dates = parsed_dates.dropna()

            if not valid_dates.empty:
                filtered_df["_parsed_datetime"] = parsed_dates
                min_date = valid_dates.min().date()
                max_date = valid_dates.max().date()

                if min_date != max_date:
                    date_range = st.sidebar.date_input(
                        "Filter Date Range",
                        value=(min_date, max_date),
                        min_value=min_date,
                        max_value=max_date
                    )

                    if isinstance(date_range, tuple) and len(date_range) == 2:
                        start_date, end_date = date_range
                        mask = (
                            (filtered_df["_parsed_datetime"].dt.date >= start_date) & 
                            (filtered_df["_parsed_datetime"].dt.date <= end_date)
                        )
                        filtered_df = filtered_df[mask]
                
                filtered_df = filtered_df.drop(columns=["_parsed_datetime"], errors="ignore")
            else:
                st.sidebar.warning("Selected column could not be parsed as valid dates.")

        st.success(f"File **{uploaded_file.name}** processed successfully!")

        # -----------------------------------------------------------------------------
        # OVERVIEW METRICS
        # -----------------------------------------------------------------------------
        total_rows = len(filtered_df)
        total_net_power = filtered_df["WHRS NET KW RECOVERED PER DAY"].sum()
        total_roi = filtered_df["Daily ROI (USD)"].sum()

        col1, col2, col3 = st.columns(3)
        col1.metric("Filtered Records", f"{total_rows}")
        col2.metric("Total Net Energy Recovered", f"{total_net_power:,.1f} kWh")
        col3.metric("Total Financial ROI Savings", f"${total_roi:,.2f} USD")

        # -----------------------------------------------------------------------------
        # VISUALIZATIONS
        # -----------------------------------------------------------------------------
        st.divider()
        st.subheader("📈 Visualizations & Trends")

        tab1, tab2 = st.tabs(["⚡ Trends Over Time", "⚓ Engine Load / RH vs Net Energy"])

        with tab1:
            st.markdown("#### Recovered Energy & Financial ROI Progression")
            
            # Dual-axis chart for Energy and ROI
            fig_trends = go.Figure()

            x_axis = filtered_df[selected_dt_col] if selected_dt_col else filtered_df.index

            # Net Energy Bar
            fig_trends.add_trace(go.Bar(
                x=x_axis,
                y=filtered_df["WHRS NET KW RECOVERED PER DAY"],
                name="Net kWh Recovered / Day",
                marker_color="#1f77b4",
                opacity=0.7
            ))

            # Gross Energy Line
            fig_trends.add_trace(go.Scatter(
                x=x_axis,
                y=filtered_df["WHRS KW RECOVERED PER DAY"],
                name="Gross kWh Recovered / Day",
                mode="lines+markers",
                line=dict(color="#ff7f0e", width=2)
            ))

            # Cumulative ROI Line (Secondary Y-Axis)
            fig_trends.add_trace(go.Scatter(
                x=x_axis,
                y=filtered_df["Total ROI (USD)"],
                name="Cumulative ROI ($)",
                mode="lines",
                line=dict(color="#2ca02c", width=3, dash="dot"),
                yaxis="y2"
            ))

            fig_trends.update_layout(
                title="Daily Energy Recovery vs Cumulative Financial Savings",
                xaxis_title="Date / Log Entry",
                yaxis=dict(title="Energy Recovered (kWh)"),
                yaxis2=dict(
                    title="Cumulative ROI (USD)",
                    overlaying="y",
                    side="right"
                ),
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                hovermode="x unified",
                template="plotly_white",
                height=480
            )

            st.plotly_chart(fig_trends, use_container_width=True)

        with tab2:
            st.markdown("#### Engine Load / Running Hours vs. Net Energy Recovered")

            # Detect potential Engine Load or Running Hours columns
            numeric_cols = filtered_df.select_dtypes(include=[np.number]).columns.tolist()
            potential_load_cols = [c for c in numeric_cols if any(k in c.upper() for k in ["LOAD", "ME KW", "POWER", "RH", "SPEED", "RPM"])]
            
            default_x_idx = numeric_cols.index(potential_load_cols[0]) if potential_load_cols else 0

            load_col = st.selectbox(
                "Select Load / Engine Variable (X-Axis):",
                options=numeric_cols,
                index=default_x_idx
            )

            if load_col:
                fig_scatter = px.scatter(
                    filtered_df,
                    x=load_col,
                    y="WHRS NET KW RECOVERED PER DAY",
                    color="Gen SFOC (g/kWh)",
                    size="WHRS KWh",
                    hover_data=[selected_dt_col, "Fuel Price ($/mt)", "Daily ROI (USD)"] if selected_dt_col else ["Daily ROI (USD)"],
                    trendline="ols",
                    title=f"{load_col} vs Net kWh Recovered",
                    labels={
                        load_col: f"{load_col}",
                        "WHRS NET KW RECOVERED PER DAY": "Net Energy Recovered (kWh)"
                    },
                    template="plotly_white"
                )

                fig_scatter.update_layout(height=480)
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

        # 1. Export as Excel (.xlsx)
        buffer_excel = io.BytesIO()
        with pd.ExcelWriter(buffer_excel, engine='openpyxl') as writer:
            filtered_df.to_excel(writer, index=False, sheet_name='WHRS_Calculated_ROI')
        buffer_excel.seek(0)

        exp_col1.download_button(
            label="📄 Download Clean Excel (.xlsx)",
            data=buffer_excel,
            file_name="WHRS_LOG_Calculated_ROI.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # 2. Export as CSV (.csv)
        buffer_csv = filtered_df.to_csv(index=False).encode('utf-8')
        exp_col2.download_button(
            label="📑 Download Clean CSV (.csv)",
            data=buffer_csv,
            file_name="WHRS_LOG_Calculated_ROI.csv",
            mime="text/csv"
        )

    except Exception as e:
        st.error(f"Error processing the uploaded file: {str(e)}")

else:
    st.info("👆 Please upload your WHRS Excel log file or CSV to execute calculations and view analytical trends.")