import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import os
import tkinter as tk
from tkinter import filedialog
import tempfile


# -------------------------------------------------
# ✅ Tkinter Folder Picker
# -------------------------------------------------
def pick_folder():
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    folder = filedialog.askdirectory()
    root.destroy()
    return folder


# -------------------------------------------------
# ✅ Conversion Function
# -------------------------------------------------
def convert_single_file(input_file_path, output_folder):

    ext = os.path.splitext(input_file_path)[1].lower()

    # ✅ Explicit Excel engines
    if ext == ".xlsx":
        sap = pd.read_excel(input_file_path, engine="openpyxl")
    elif ext == ".xls":
        sap = pd.read_excel(input_file_path, engine="xlrd")
    else:
        raise ValueError("Unsupported Excel format")

    # Cleaning
    sap_no_na = sap.dropna().reset_index(drop=True)
    sap_no_na = sap_no_na[abs(sap_no_na.iloc[:, 4]) != 0]
    dim_df = sap_no_na.shape

    jp = pd.DataFrame(np.nan, index=range(dim_df[0] + 1), columns=range(28))

    # Header
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

    # Zero-fill
    jp.iloc[1:1+n, 6] = (
        sap_no_na.iloc[:n, 3].astype(str).str.split('.').str[0].str.zfill(8)
    )

    jp.iloc[1:1+n, 7] = 0
    jp.iloc[1:1+n, 8] = (sap_no_na.iloc[:n, 4] * 100).abs().round().astype(int)

    # constants
    jp.iloc[1:1+n, 9] = "GBP"
    jp.iloc[1:1+n, 10] = "GIR"
    jp.iloc[1:1+n, 11] = "01"
    jp.iloc[1:1+n, 12] = 99
    jp.iloc[1:1+n, 13] = 76904122

    # Truncate name
    jp.iloc[1:1+n, 16] = sap_no_na.iloc[:n, 1].astype(str).str.slice(0, 18)

    # Footer
    jp.iloc[-1, 0] = "FT"
    jp.iloc[-1, 1] = dim_df[0]
    jp.iloc[-1, 2] = jp.shape[0]
    jp.iloc[-1, 3] = (sap_no_na.iloc[:n, 4] * 100).abs().round().astype(int).sum()

    base_name = os.path.splitext(os.path.basename(input_file_path))[0]

    # Save output
    jp.to_excel(os.path.join(output_folder, f"{base_name}.xlsx"), index=False, header=False)
    jp.to_csv(os.path.join(output_folder, f"{base_name}.csv"), index=False, header=False)
    jp.to_csv(os.path.join(output_folder, f"{base_name}.txt"), index=False, header=False, sep=",", quoting=0)


# -------------------------------------------------
# ✅ Streamlit UI
# -------------------------------------------------
st.title("📄 SAP → JPMorgan Batch Converter (Local Output Folder)")
st.write("Upload SAP files → choose output folder → convert.")


uploaded_files = st.file_uploader(
    "Upload SAP Excel files:",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

# Store folder in session
if "output_folder" not in st.session_state:
    st.session_state.output_folder = ""


# ✅ Folder picker button
if st.button("📁 Choose Output Folder"):
    folder = pick_folder()
    if folder:
        st.session_state.output_folder = folder

st.write("✅ Selected Output Folder:", st.session_state.output_folder)


# -------------------------------------------------
# ✅ Start Conversion
# -------------------------------------------------
if st.button("Start Conversion"):

    if not uploaded_files:
        st.error("❌ Please upload SAP files.")
        st.stop()

    if not st.session_state.output_folder:
        st.error("❌ Please choose an output folder.")
        st.stop()

    temp_dir = tempfile.gettempdir()

    # ✅ Store file bytes ONCE
    file_map = {}  # filename → bytes
    for uf in uploaded_files:
        file_map[uf.name] = uf.read()

    # ✅ Determine input dirs
    input_dirs = []

    for name, data in file_map.items():

        temp_path = os.path.join(temp_dir, name)

        with open(temp_path, "wb") as f:
            f.write(data)

        input_dirs.append(os.path.abspath(os.path.dirname(temp_path)))

    # ✅ Prevent output = input folder
    if any(os.path.abspath(st.session_state.output_folder) == d for d in input_dirs):
        st.error("⚠ Output folder cannot be the same as input folder.")
        st.stop()

    # ✅ Convert files (reusing cached bytes)
    try:
        for name, data in file_map.items():

            temp_path = os.path.join(temp_dir, name)

            with open(temp_path, "wb") as f:
                f.write(data)

            convert_single_file(temp_path, st.session_state.output_folder)

        st.success("✅ All files converted successfully!")
        st.balloons()

    except Exception as e:
        st.error(f"❌ Error: {e}")