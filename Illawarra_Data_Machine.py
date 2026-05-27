import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from openpyxl import load_workbook

st.set_page_config(
    page_title="Illawarra Data Machine",
    page_icon="🌊",
    layout="wide"
)

st.markdown(
    """
    <style>
    .block-container {padding-top: 2rem; padding-bottom: 2rem;}
    .idm-card {
        background: #f7f9fc;
        border: 1px solid #dbe4f0;
        border-radius: 14px;
        padding: 1rem 1.2rem;
        margin-bottom: 1rem;
    }
    .idm-small {color: #5b6573; font-size: 0.95rem;}
    .idm-title {font-size: 2rem; font-weight: 700; margin-bottom: 0.25rem;}
    .idm-pill {
        display: inline-block;
        background: #e8f2ff;
        color: #0b5cab;
        padding: 0.2rem 0.6rem;
        border-radius: 999px;
        font-size: 0.85rem;
        font-weight: 600;
        margin-bottom: 0.6rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


def import_beach_watch_xlsx(uploaded_file):
    uploaded_file.seek(0)
    wb = load_workbook(uploaded_file, read_only=True, data_only=True)
    ws = wb.active

    header_row = None
    for i, row in enumerate(ws.iter_rows(values_only=True), start=1):
        if row and any(str(x).strip() == "Sample Date" for x in row if x is not None):
            header_row = i
            break

    if header_row is None:
        raise ValueError("Could not find header row containing 'Sample Date'")

    uploaded_file.seek(0)
    df = pd.read_excel(uploaded_file, sheet_name=ws.title, header=header_row - 1, engine="openpyxl")
    df = df.dropna(axis=1, how="all").dropna(how="all").reset_index(drop=True)
    df.columns = [str(c).strip() for c in df.columns]
    return df


def clean_beach_watch_df(df):
    df = df.copy()

    for col in ["Sample Run", "Sample Run Submission", "Site Sample Run", "Site", "Sampler", "Status"]:
        if col in df.columns:
            df[col] = df[col].astype("object")

    rename_map = {
        "Time": "Sample Time",
        "Field Conductivity (mS/cm)": "Conductivity (µS/cm)",
        "Dissolved oxygen (mg/L)": "Dissolved oxygen (mg/L)",
        "No Of Swimmers": "Number of swimmers",
        "Blue Bottles": "Bluebottles",
        "Leaf Litter": "Leaf litter",
        "Marine Debris": "Debris",
        "Surface Foam Scum": "Surface scum",
        "Weeds": "Weed",
        "Enterococci Result (CFU/100mL)": "Enterococci Result",
        "Water Temperature (0C)": "Water Temperature (℃)",
    }
    df = df.rename(columns=rename_map)

    for col in df.columns:
        if df[col].dtype == "object":
            df[col] = df[col].replace(r"^\s*$", np.nan, regex=True)

    key_cols = [c for c in ["Sample Date", "Site Sample Run", "Site", "Sampler"] if c in df.columns]
    if key_cols:
        df = df.dropna(subset=key_cols, how="all").copy()

    if "Sample Date" in df.columns:
        sample_date_dt = pd.to_datetime(df["Sample Date"], errors="coerce", dayfirst=True)
        df["Sample Date"] = sample_date_dt.apply(
            lambda x: f"{x.day}/{x.month}/{x.year}" if pd.notna(x) else np.nan
        )

    if "Sample Time" in df.columns:
        t = pd.to_datetime(df["Sample Time"], errors="coerce")
        df["Sample Time"] = t.dt.strftime("%H:%M:%S") + ".000Z"
        df.loc[t.isna(), "Sample Time"] = np.nan

    numeric_cols = [
        "Water Temperature (℃)",
        "Conductivity (µS/cm)",
        "Dissolved oxygen (mg/L)",
        "Number of swimmers",
        "Enterococci Result",
        "Drain Flow",
        "Lagoon Flow",
        "Surface scum",
        "Leaf litter",
        "Litter",
        "Debris",
        "Bluebottles",
        "Weed",
        "Esky Temperature (℃)",
    ]
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    if "Conductivity (µS/cm)" in df.columns:
        df["Conductivity (µS/cm)"] = df["Conductivity (µS/cm)"] * 1000

    real_sample_mask = pd.Series(False, index=df.index)
    for c in ["Site", "Site Sample Run", "Sample Date", "Sampler"]:
        if c in df.columns:
            real_sample_mask = real_sample_mask | df[c].notna()

    df["Sample Run"] = pd.Series(np.nan, index=df.index, dtype="object")
    df.loc[real_sample_mask, "Sample Run"] = "Illawarra Beaches"

    df["Sample Run Submission"] = pd.Series(np.nan, index=df.index, dtype="object")
    if "Sample Date" in df.columns:
        valid_submission = real_sample_mask & df["Sample Date"].notna()
        df.loc[valid_submission, "Sample Run Submission"] = (
            "Illawarra Beaches-"
            + df.loc[valid_submission, "Sample Date"].astype(str)
        )

    df["Status"] = pd.Series(np.nan, index=df.index, dtype="object")
    real_row_mask = pd.Series(False, index=df.index)
    for c in ["Sample Run", "Sample Date", "Site", "Site Sample Run"]:
        if c in df.columns:
            real_row_mask = real_row_mask | df[c].notna()
    df.loc[real_row_mask, "Status"] = "Awaiting Approval"

    df["No Enterococci Result"] = np.nan
    if "Enterococci Result" in df.columns:
        real_sample_mask = pd.Series(False, index=df.index)
        for c in ["Site", "Site Sample Run", "Sample Run", "Sample Date"]:
            if c in df.columns:
                real_sample_mask = real_sample_mask | df[c].notna()
        no_enterococci_mask = df["Enterococci Result"].isna() & real_sample_mask
        df.loc[no_enterococci_mask, "No Enterococci Result"] = 1

    desired_order = [
        "Sample Run",
        "Sample Run Submission",
        "Sample Date",
        "Sampler",
        "Observation Type",
        "Site Sample Run",
        "Site",
        "Other Site Name",
        "Sample Time",
        "Weather",
        "Drain Flow",
        "Lagoon Flow",
        "Water Temperature (℃)",
        "Conductivity (µS/cm)",
        "Dissolved oxygen (mg/L)",
        "Number of swimmers",
        "Surface scum",
        "Leaf litter",
        "Litter",
        "Debris",
        "Bluebottles",
        "Weed",
        "Visual Turbidity",
        "Esky Temperature (℃)",
        "Comments",
        "Enterococci Result",
        "No Enterococci Result",
        "Status",
    ]

    for col in desired_order:
        if col not in df.columns:
            if col in ["Sample Run", "Sample Run Submission", "Status"]:
                df[col] = pd.Series(np.nan, index=df.index, dtype="object")
            else:
                df[col] = np.nan

    export_key_cols = [
        "Sample Run Submission",
        "Sample Date",
        "Site Sample Run",
        "Site",
        "Sample Time",
    ]

    existing_export_key_cols = [col for col in export_key_cols if col in df.columns]
    if existing_export_key_cols:
        df = df.dropna(subset=existing_export_key_cols, how="all").copy()

    return df[desired_order].reset_index(drop=True)


def get_output_filename(df):
    if "Sample Date" not in df.columns:
        raise ValueError("Sample Date column not found")

    sample_dates = pd.to_datetime(df["Sample Date"], errors="coerce", dayfirst=True)
    valid_dates = sample_dates.dropna()

    if valid_dates.empty:
        raise ValueError("No valid dates found in Sample Date column")

    file_date = valid_dates.min().strftime("%d %B %Y")
    return f"{file_date} Illawarra Beaches Data.csv"


def dataframe_to_csv_bytes(df):
    return df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")


st.markdown('<div class="idm-title"> P.O.Os Illawarra Data Machine</div>', unsafe_allow_html=True)
st.markdown('<div class="idm-small">Automatically cleans Illawarra data, ready for Salesforce Upload.</div>', unsafe_allow_html=True)

with st.sidebar:
    st.header("How it works")
    st.markdown(
        """
1. Upload the raw Illawarra xlsx file.
2. Check that everything looks right in the preview.
3. Hit download, then import that csv into Salesforce and attach it to the run.
        """
    )
    st.info("""
For Salesforce imports  
1. Go to Submissions  
2. Add new records  
3. Change the 4th box down to Site Sample Run Name  
4. Skip 2 boxes  
5. Change the next 3 boxes to Site name, Contact name & Sample Run Submission Name  
6. Finally add the file and change the encoder to UTF-8
    """)

uploaded_file = st.file_uploader("Upload Beach Watch Excel file", type=["xlsx"])

if uploaded_file is None:
    st.markdown('<div class="idm-card"><strong>Start here:</strong> upload an Excel file above to generate the cleaned CSV.</div>', unsafe_allow_html=True)
else:
    try:
        with st.spinner("Importing and cleaning file..."):
            raw_df = import_beach_watch_xlsx(uploaded_file)
            clean_df = clean_beach_watch_df(raw_df)
            output_filename = get_output_filename(clean_df)
            csv_bytes = dataframe_to_csv_bytes(clean_df)

        col1, col2, col3 = st.columns(3)
        col1.metric("Rows", f"{len(clean_df):,}")
        col2.metric("Columns", f"{len(clean_df.columns):,}")
        col3.metric("Output file", output_filename)

        st.download_button(
            label="Download cleaned CSV",
            data=csv_bytes,
            file_name=output_filename,
            mime="text/csv",
            use_container_width=True,
        )

        with st.expander("Preview cleaned data", expanded=True):
            st.dataframe(clean_df, use_container_width=True, height=420)

        with st.expander("Column names"):
            st.write(clean_df.columns.tolist())

        with st.expander("Raw data preview"):
            st.dataframe(raw_df.head(20), use_container_width=True)

    except Exception as e:
        st.error(f"Error: {e}")