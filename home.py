import streamlit as st
import pandas as pd
import zipfile
import tempfile
import os
import base64
from datetime import datetime

# Upload comma-separated CSV file(s)
uploaded_files = st.file_uploader(
    "Choose a CSV file",
    accept_multiple_files=True,
    type=['csv', 'zip'],
    )

# Option to toggle automatic type conversion (e.g. 1,000 to 1000)
auto_conversion = st.toggle(
    "Automatic type conversion",
    value=False,
    help="Sometimes, numbers are stored as text in CSV files. The delimiters and thousands separators may not work properly in that case. This option will automatically convert them to numbers. If you are not sure, leave this option disabled and check the result."
    )

# Initiate data dictionary
data = {}

# Filename function
def file_name():
    return str(datetime.now().strftime("%Y-%m-%d_%H-%M_csv2xlsx_export"))

if len(uploaded_files) > 0:
    with st.status("Display progress"):
        # Iterate through each uploaded file and check if they are CSV or ZIP
        for uploaded_file in uploaded_files:
            if uploaded_file.name.endswith('.csv'):
                if auto_conversion:
                    st.write(f"📄 Reading file: **{uploaded_file.name}**")
                    df = pd.read_csv(uploaded_file)

                    # Automatic conversion to numeric logic
                    for col in df.columns:
                        try:
                            df[col] = pd.to_numeric(df[col].str.replace(",", "").str.replace("\"", ""), errors='raise')
                            st.write(f"⚠️ Automatically converted column in **{uploaded_file.name}**: {col}")
                        except:
                            pass

                        data[uploaded_file.name] = df
                        st.write(f"💾 Data from file **{uploaded_file.name}** was successfully processed")
                else:
                    st.write(f"📄 Reading file: **{uploaded_file.name}**")
                    data[uploaded_file.name] = pd.read_csv(uploaded_file)
                    st.write(f"💾 Data from file **{uploaded_file.name}** was successfully processed")

            elif uploaded_file.name.endswith('.zip'):
                with tempfile.TemporaryDirectory() as tmpdirname:
                    with zipfile.ZipFile(uploaded_file) as z:
                        z.extractall(tmpdirname)
                        for filename in z.namelist():
                            if filename.endswith('.csv'):
                                if auto_conversion:
                                    df = pd.read_csv(os.path.join(tmpdirname, filename))

                                    # Automatic conversion to numeric logic
                                    for col in df.columns:
                                        try:
                                            df[col] = pd.to_numeric(df[col].str.replace(",", "").str.replace("\"", ""), errors='raise')
                                            st.write(f"⚠️ Automatically converted column in **{filename}**: {col}")
                                        except:
                                            pass

                                    data[filename] = df
                                else:
                                    st.write(f"📄 Reading file: **{filename}**")
                                    data[filename] = pd.read_csv(os.path.join(tmpdirname, filename))
                                    st.write(f"💾 Data from file **{filename}** was successfully processed")
                            else:
                                st.warning(f"⚠️ File **{filename}** is not a CSV file. Skipping...")

# Preview
if len(data) > 0:
    with st.expander('Preview data'):
        preview_selector = st.selectbox('Preview', list(data.keys()))
        st.dataframe(data[preview_selector])

# Download
if len(data) == 0:
    st.info('Upload CSV or ZIP file(s).')
elif len(data) > 1:
    download_col1, download_col2 = st.columns(2)

    with download_col1:
        with st.expander('Option 1: Single file', expanded=True):
            st.write("Download all data as a single XSLX file, with each CSV file as a separate sheet.")
            # Download as single file with multiple sheets
            with tempfile.TemporaryDirectory() as tmpdirname:
                with pd.ExcelWriter(os.path.join(tmpdirname, f'{file_name()}.xlsx')) as writer:
                    for key, value in data.items():
                        # Create sheet name from filename, emitting the extension
                        sheet_name = key[:-4][:31]  # Excel sheet name cannot be longer than 31 characters
                        value.to_excel(writer, sheet_name=sheet_name, index=False)
                with open(os.path.join(tmpdirname, f'{file_name()}.xlsx'), 'rb') as f:
                    bytes_data = f.read()
                b64 = base64.b64encode(bytes_data).decode()
                href = f'data:file/xlsx;base64,{b64}'
                st.download_button(label="Download", data=bytes_data, file_name=f'{file_name()}.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', use_container_width=True)

    with download_col2:
        with st.expander('Option 2: Separate files', expanded=True):
            st.write("Download data as individual XLSX files with a single sheet, put together in a ZIP file.")
            # Download as separate files in a ZIP
            with tempfile.TemporaryDirectory() as tmpdirname:
                with zipfile.ZipFile(os.path.join(tmpdirname, f'{file_name()}.zip'), 'w') as z:
                    for key, value in data.items():
                        xlsx_filename = f'{key[:-4]}.xlsx'
                        value.to_excel(os.path.join(tmpdirname, xlsx_filename), index=False)
                        z.write(os.path.join(tmpdirname, xlsx_filename), arcname=xlsx_filename)
                with open(os.path.join(tmpdirname, f'{file_name()}.zip'), 'rb') as f:
                    bytes_data = f.read()
                b64 = base64.b64encode(bytes_data).decode()
                href = f'data:file/zip;base64,{b64}'
                st.download_button(label="Download", data=bytes_data, file_name=f'{file_name()}.zip', mime='application/zip', use_container_width=True)

else:
    # Download
    with tempfile.TemporaryDirectory() as tmpdirname:
        single_filename = list(data.keys())[0][:-4] + '.xlsx'
        data[list(data.keys())[0]].to_excel(os.path.join(tmpdirname, single_filename), index=False)
        with open(os.path.join(tmpdirname, single_filename), 'rb') as f:
            bytes_data = f.read()
        b64 = base64.b64encode(bytes_data).decode()
        href = f'data:file/xlsx;base64,{b64}'
        st.download_button(label="Download", data=bytes_data, file_name=single_filename, mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', use_container_width=True)