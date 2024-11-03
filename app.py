import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
from openpyxl import load_workbook

def load_data(uploaded_files):
    all_data = pd.DataFrame()
    required_columns = ['ALLOCATED TO', 'STATUS', 'PRODUCT_DESCRIPTION', 'DATE', 'File Name']

    for uploaded_file in uploaded_files:
        if uploaded_file.name.endswith(".xlsx"):
            workbook = load_workbook(uploaded_file, data_only=True)
            sheet = workbook.active

            # Extract data including hidden columns
            data = pd.DataFrame(sheet.values)

            # Set the first row as the header if not already set
            data.columns = data.iloc[0]
            data = data[1:]  # Remove the header row

            # Remove duplicate columns
            data = data.loc[:, ~data.columns.duplicated()]

            # Standardize column names to avoid casing issues and trailing spaces
            data.columns = data.columns.str.strip().str.upper()

            # Ensure each required column exists, or create it with None values
            for col in required_columns:
                if col.upper() not in data.columns:
                    data[col] = None

            # Set 'File Name' column based on the uploaded file name if it’s missing
            if 'File Name' not in data.columns:
                data['File Name'] = uploaded_file.name

            # Select only the required columns to ensure consistent structure
            data = data.reindex(columns=required_columns)

            # Append to the main DataFrame
            all_data = pd.concat([all_data, data], ignore_index=True)

    return all_data

def main():
    st.title("Live Data Tracking Dashboard")
    uploaded_files = st.file_uploader("Choose Excel files", type="xlsx", accept_multiple_files=True)

    if st.button("Load Files") and uploaded_files:
        all_data = load_data(uploaded_files)

        if not all_data.empty:
            st.success("Data loaded successfully!")
            # Additional processing and visualization code...
        else:
            st.warning("No data found in the uploaded files.")
    else:
        st.info("Please upload Excel files to get started.")

if __name__ == "__main__":
    main()
