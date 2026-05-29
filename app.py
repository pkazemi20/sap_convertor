import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import io
import zipfile


def login():
    st.title("🔐 Login")

    username = st.text_input("Username")
    password = st.text_input("Password", type="password")

    if st.button("Login"):
        if username in st.secrets["USERS"]:
            idx = st.secrets["USERS"].index(username)
            if password == st.secrets["PASSWORDS"][idx]:
                st.session_state.authenticated = True
                st.rerun()
        st.error("❌ Invalid username or password")


if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    login()
    st.stop()


# -------------------------------------------------
# Conversion Function
# -------------------------------------------------
def convert_single_file(uploaded_file, currency):

    ext = uploaded_file.name.split(".")[-1].lower()

    # Read Excel as text to preserve leading zeros
    if ext == "xlsx":
        sap = pd.read_excel(
            uploaded_file,
            engine="openpyxl",
            dtype=str
        )
    elif ext == "xls":
        sap = pd.read_excel(
            uploaded_file,
            engine="xlrd",
            dtype=str
        )
    else:
        raise ValueError("Unsupported file type")

    sap_no_na = sap.dropna().reset_index(drop=True)

    # Filter out zero amounts
    amount_col = pd.to_numeric(
        sap_no_na.iloc[:, 4],
        errors="coerce"
    ).fillna(0)

    sap_no_na = sap_no_na[
        amount_col.abs() != 0
    ].reset_index(drop=True)

    dim_df = sap_no_na.shape

    jp = pd.DataFrame(
        np.nan,
        index=range(dim_df[0] + 1),
        columns=range(28)
    )

    # Header row
    header_row = [
        "FH",
        "HVLPharma",
        datetime.now().strftime("%Y%m%d"),
        datetime.now().strftime("%H%M%S"),
        "01100"
    ]

    jp.loc[-1] = header_row + [None] * (
        jp.shape[1] - len(header_row)
    )

    jp = jp.sort_index().reset_index(drop=True)

    n = min(len(jp) - 1, len(sap_no_na))

    # Amounts for calculations
    amounts = pd.to_numeric(
        sap_no_na.iloc[:n, 4],
        errors="coerce"
    ).fillna(0)

    # -------------------------
    # Data rows
    # -------------------------

    jp.iloc[1:, 0] = "TR"

    # Payment document no.
    jp.iloc[1:1+n, 1] = (
        sap_no_na.iloc[:n, 0]
        .astype(str)
        .str.strip()
        .values
    )

    # Due date
    jp.iloc[1:1+n, 2] = pd.to_datetime(
        sap_no_na.iloc[:n, 5]
    ).dt.strftime("%Y%m%d")

    # Bank country
    jp.iloc[1:1+n, 3] = (
        sap_no_na.iloc[:n, 6]
        .astype(str)
        .str.strip()
        .values
    )

    # Bank number (preserve leading zeros)
    jp.iloc[1:1+n, 5] = (
        sap_no_na.iloc[:n, 2]
        .astype(str)
        .str.strip()
        .values
    )

    # Payee bank account number (preserve leading zeros)
    jp.iloc[1:1+n, 6] = (
        sap_no_na.iloc[:n, 3]
        .astype(str)
        .str.strip()
        .values
    )

    jp.iloc[1:1+n, 7] = 0

    # Amount
    jp.iloc[1:1+n, 8] = (
        amounts * 100
    ).abs().round().astype(int)

    # Currency selected by user
    jp.iloc[1:1+n, 9] = currency

    # Constants
    jp.iloc[1:1+n, 10] = "GIR"
    jp.iloc[1:1+n, 11] = "1"
    jp.iloc[1:1+n, 12] = 99
    jp.iloc[1:1+n, 13] = 76904122

    # Payee name truncated to 18 chars
    jp.iloc[1:1+n, 16] = (
        sap_no_na.iloc[:n, 1]
        .astype(str)
        .str.slice(0, 18)
    )

    # -------------------------
    # Footer row
    # -------------------------

    jp.iloc[-1, 0] = "FT"
    jp.iloc[-1, 1] = dim_df[0]
    jp.iloc[-1, 2] = jp.shape[0]

    jp.iloc[-1, 3] = (
        amounts * 100
    ).abs().round().astype(int).sum()

    return jp


# -------------------------------------------------
# Streamlit UI
# -------------------------------------------------

st.title("📄 SAP to JPMorgan File Converter")

currency = st.selectbox(
    "Select Currency",
    [
        "GBP",
        "EUR",
        "USD",
        "CHF",
        "JPY",
        "AUD",
        "CAD"
    ],
    index=0
)

uploaded_files = st.file_uploader(
    "Upload SAP files (.xls or .xlsx)",
    accept_multiple_files=True,
    type=["xls", "xlsx"]
)

if uploaded_files:

    if st.button("Convert All Files"):

        zip_buffer = io.BytesIO()

        with zipfile.ZipFile(
            zip_buffer,
            "w",
            zipfile.ZIP_DEFLATED
        ) as zipf:

            for uf in uploaded_files:

                df = convert_single_file(
                    uf,
                    currency
                )

                base = uf.name.rsplit(".", 1)[0]

                # Excel output
                xlsx_bytes = io.BytesIO()

                with pd.ExcelWriter(
                    xlsx_bytes,
                    engine="openpyxl"
                ) as writer:

                    df.to_excel(
                        writer,
                        index=False,
                        header=False
                    )

                    # Force Excel to keep text formatting
                    ws = writer.sheets["Sheet1"]

                    for row in ws.iter_rows():
                        for cell in row:
                            cell.number_format = "@"

                zipf.writestr(
                    f"{base}.xlsx",
                    xlsx_bytes.getvalue()
                )

                # CSV output
                csv_data = df.to_csv(
                    index=False,
                    header=False
                )

                zipf.writestr(
                    f"{base}.csv",
                    csv_data
                )

                # TXT output
                txt_data = df.to_csv(
                    index=False,
                    header=False,
                    sep=","
                )

                zipf.writestr(
                    f"{base}.txt",
                    txt_data
                )

        st.success(
            "✅ Conversion complete! Download your ZIP file below."
        )

        st.download_button(
            label="⬇ Download ZIP",
            data=zip_buffer.getvalue(),
            file_name="converted_files.zip",
            mime="application/zip"
        )
