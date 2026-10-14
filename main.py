import streamlit as st
import pandas as pd
import numpy as np
import io

# Page setup
st.set_page_config(
    page_title="WHRS Batch Calculator & ROI Processor",
    page_icon="⚓",
    layout="wide"
)

st.title("⚓ WHRS Log Batch File Processor & ROI Calculator")
st.markdown("""
Upload your WHRS Excel log file or CSV containing the raw input columns. 
The application will automatically initialize all required columns, run all Excel-equivalent calculations, omit duplicate columns, and generate the clean final dataset for export.
""")

# Standard fuel lookup tables from the sheet formulas
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
    fuel_series = df[fuel_col].astype(str).str.strip().str.upper()

    # 1. Gen SFOC (g/kWh) Formula:
    # =IF(TRIM(C4)="", "", IF(TRIM(C4)="LNG", 192.1, 182.4))
    df["Gen SFOC (g/kWh)"] = np.where(fuel_series == "LNG", 192.1, 182.4)

    # 2. Fuel Price ($/mt) Formula:
    # =IFS(C4="LNG", 1345, C4="LSMGO", 1280, C4="VLSFO", 820, C4="HSFO", 635)
    df["Fuel Price ($/mt)"] = fuel_series.map(FUEL_PRICE_MAP).fillna(820.0)

    # Convert numeric columns to float
    df[net_total_col] = pd.to_numeric(df[net_total_col], errors='coerce').fillna(0)
    df[gross_total_col] = pd.to_numeric(df[gross_total_col], errors='coerce').fillna(0)
    df[rh_col] = pd.to_numeric(df[rh_col], errors='coerce').fillna(0)

    # Calculate difference metrics
    net_diff = df[net_total_col].diff()
    net_diff.iloc[0] = df[net_total_col].iloc[0]
    calculated_net_recovered = net_diff.clip(lower=0)

    gross_diff = df[gross_total_col].diff()
    gross_diff.iloc[0] = df[gross_total_col].iloc[0]
    calculated_gross_recovered = gross_diff.clip(lower=0)

    # Prevent duplicate columns by dropping any pre-existing duplicate columns with matching names
    cols_to_clean = ["WHRS NET KW RECOVERED PER DAY", "WHRS  NET KW RECOVERED PER DAY", "WHRS KW RECOVERED PER DAY"]
    df = df.loc[:, ~df.columns.isin(cols_to_clean)]

    # Assign calculated metrics
    df["WHRS NET KW RECOVERED PER DAY"] = calculated_net_recovered
    df["WHRS KW RECOVERED PER DAY"] = calculated_gross_recovered

    # 5. WHRS KWh (Average power generation per running hour)
    # Excel Formula: IF(AND([@[WHRS KW RECOVERED PER DAY]]<>0, [@[WHRS RH]]<>0), ROUND([@[WHRS KW RECOVERED PER DAY]]/[@[WHRS RH]], 1), 0)
    
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

    # 6. Daily ROI (USD) Formula:
    df["Daily ROI (USD)"] = (
        df["WHRS NET KW RECOVERED PER DAY"] * 
        (df["Gen SFOC (g/kWh)"] / 1000000.0) * 
        df["Fuel Price ($/mt)"]
    )

    # 7. Total ROI (USD) Formula:
    df["Total ROI (USD)"] = df["Daily ROI (USD)"].cumsum()

    # Remove duplicate columns if present in raw sheet structure
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
                if "ME FUEL IN USE" in row_str or "DATE TIME" in row_str or "VESSEL STATUS" in row_str:
                    header_row = idx
                    break
            
            raw_df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=header_row)

        st.success(f"File **{uploaded_file.name}** loaded successfully! Processing calculations...")

        # Run automated calculations
        computed_df = compute_whrs_metrics(raw_df)

        # Overview Metrics
        total_days = len(computed_df)
        total_net_power = computed_df["WHRS NET KW RECOVERED PER DAY"].sum()
        total_roi = computed_df["Total ROI (USD)"].iloc[-1] if not computed_df.empty else 0.0

        col1, col2, col3 = st.columns(3)
        col1.metric("Total Rows Processed", f"{total_days}")
        col2.metric("Total Net Energy Recovered", f"{total_net_power:,.1f} kWh")
        col3.metric("Total Financial ROI Savings", f"${total_roi:,.2f} USD")

        st.divider()
        st.subheader("📊 Calculated Dataset Preview")
        
        st.dataframe(computed_df, use_container_width=True)

        # -----------------------------------------------------------------------------
        # EXPORT OPTIONS
        # -----------------------------------------------------------------------------
        st.divider()
        st.subheader("📥 Export Final Processed File")
        
        exp_col1, exp_col2 = st.columns(2)

        # 1. Export as Excel (.xlsx)
        buffer_excel = io.BytesIO()
        with pd.ExcelWriter(buffer_excel, engine='openpyxl') as writer:
            computed_df.to_excel(writer, index=False, sheet_name='WHRS_Calculated_ROI')
        buffer_excel.seek(0)

        exp_col1.download_button(
            label="📄 Download Clean Excel (.xlsx)",
            data=buffer_excel,
            file_name="WHRS_LOG_Calculated_ROI.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # 2. Export as CSV (.csv)
        buffer_csv = computed_df.to_csv(index=False).encode('utf-8')
        exp_col2.download_button(
            label="📑 Download Clean CSV (.csv)",
            data=buffer_csv,
            file_name="WHRS_LOG_Calculated_ROI.csv",
            mime="text/csv"
        )

    except Exception as e:
        st.error(f"Error processing the uploaded file: {str(e)}")

else:
    st.info("👆 Please upload your WHRS Excel log file or CSV to execute the calculations.")