import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
from openpyxl import load_workbook

def load_data(uploaded_files):
    all_data = pd.DataFrame()
    required_columns = ['ALLOCATED TO', 'STATUS', 'PRODUCT_DESCRIPTION', 'DATE', 'File Name']

    for uploaded_file in uploaded_files:
        if uploaded_file.name.endswith(".xlsx"):
            # Load the workbook with openpyxl
            workbook = load_workbook(uploaded_file, data_only=True)
            sheet = workbook.active

            # Extract all data, including hidden columns
            data = pd.DataFrame(sheet.values)

            # Set first row as header
            data.columns = data.iloc[0]
            data = data[1:]

            # Standardize column names by stripping whitespace and converting to uppercase
            data.columns = data.columns.str.strip().str.upper()

            # Rename columns to match required columns if they have slight variations
            column_mapping = {
                'ALLOCATED TO': 'ALLOCATED TO', 
                'STATUS': 'STATUS', 
                'PRODUCT_DESCRIPTION': 'PRODUCT_DESCRIPTION', 
                'DATE': 'DATE',
                'FILE NAME': 'File Name'
            }
            data.rename(columns=column_mapping, inplace=True)

            # If 'File Name' column is missing, add it
            if 'File Name' not in data.columns:
                data['File Name'] = uploaded_file.name

            # Ensure all required columns are present in the DataFrame
            for col in required_columns:
                if col not in data.columns:
                    data[col] = None

            # Append data to the combined DataFrame
            all_data = pd.concat([all_data, data[required_columns]], ignore_index=True)

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
