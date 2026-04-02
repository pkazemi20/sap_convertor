import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import io
import zipfile


# -------------------------------------------------
# ✅ Conversion Function (same transformation logic as before)
# -------------------------------------------------
def convert_single_file(uploaded_file):

    ext = uploaded_file.name.split(".")[-1].lower()

    # Read Excel
    if ext == "xlsx":
        sap = pd.read_excel(uploaded_file, engine="openpyxl")
    elif ext == "xls":
        sap = pd.read_excel(uploaded_file, engine="xlrd")
    else:
        raise ValueError("Unsupported file type")

    sap_no_na = sap.dropna().reset_index(drop=True)
    sap_no_na = sap_no_na[abs(sap_no_na.iloc[:, 4]) != 0]
    dim_df = sap_no_na.shape

    jp = pd.DataFrame(np.nan, index=range(dim_df[0] + 1), columns=range(28))

    # Header row
    header_row = [
        "FH", "HVLPharma",
        datetime.now().strftime("%Y%m%d"),
        datetime.now().strftime("%H%M%S"),
        "01100"
    ]
    jp.loc[-1] = header_row + [None] * (jp.shape[1] - len(header_row))
    jp = jp.sort_index().reset_index(drop=True)

    n = min(len(jp) - 1, len(sap_no_na))

    # Data rows
    jp.iloc[1:, 0] = "TR"
    jp.iloc[1:1+n, 1] = sap_no_na.iloc[:n, 0].values
    jp.iloc[1:1+n, 2] = pd.to_datetime(sap_no_na.iloc[:n, 5]).dt.strftime('%Y%m%d')
    jp.iloc[1:1+n, 3] = sap_no_na.iloc[:n, 6].values
    jp.iloc[1:1+n, 5] = sap_no_na.iloc[:n, 2].astype(int).values
    jp.iloc[1:1+n, 6] = (
        sap_no_na.iloc[:n, 3].astype(str).str.split('.').str[0].str.zfill(8)
    )
    jp.iloc[1:1+n, 7] = 0
    jp.iloc[1:1+n, 8] = (sap_no_na.iloc[:n, 4] * 100).abs().round().astype(int)

    # Constant fields
    jp.iloc[1:1+n, 9] = "GBP"
    jp.iloc[1:1+n, 10] = "GIR"
    jp.iloc[1:1+n, 11] = "01"
    jp.iloc[1:1+n, 12] = 99
    jp.iloc[1:1+n, 13] = 76904122

    # Truncate name
    jp.iloc[1:1+n, 16] = sap_no_na.iloc[:n, 1].astype(str).str.slice(0, 18)

    # Footer row
    jp.iloc[-1, 0] = "FT"
    jp.iloc[-1, 1] = dim_df[0]
    jp.iloc[-1, 2] = jp.shape[0]
    jp.iloc[-1, 3] = (sap_no_na.iloc[:n, 4] * 100).abs().round().astype(int).sum()

    return jp


# -------------------------------------------------
# ✅ Streamlit UI
# -------------------------------------------------
st.title("📄 SAP to JPMorgan File Converter")

uploaded_files = st.file_uploader(
    "Upload SAP files (.xls or .xlsx)",
    accept_multiple_files=True,
    type=["xls", "xlsx"]
)

if uploaded_files:
    if st.button("Convert All Files"):

        zip_buffer = io.BytesIO()

        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:

            for uf in uploaded_files:
                df = convert_single_file(uf)
                base = uf.name.rsplit(".", 1)[0]

                # Excel
                xlsx_bytes = io.BytesIO()
                df.to_excel(xlsx_bytes, index=False, header=False)
                zipf.writestr(f"{base}.xlsx", xlsx_bytes.getvalue())

                # CSV
                csv_data = df.to_csv(index=False, header=False)
                zipf.writestr(f"{base}.csv", csv_data)

                # TXT
                txt_data = df.to_csv(index=False, header=False, sep=",")
                zipf.writestr(f"{base}.txt", txt_data)

        st.success("✅ Conversion complete! Download your ZIP file below.")

        st.download_button(
            label="⬇ Download ZIP",
            data=zip_buffer.getvalue(),
            file_name="converted_files.zip",
            mime="application/zip"
        )
