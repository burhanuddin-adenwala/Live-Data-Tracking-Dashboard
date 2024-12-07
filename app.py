import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
from openpyxl import load_workbook
import zipfile
import os
import tempfile


def extract_files_from_zip(uploaded_files):
    """Extracts Excel files from uploaded zip files."""
    extracted_files = []
    for uploaded_file in uploaded_files:
        if uploaded_file.name.endswith(".zip"):
            with tempfile.TemporaryDirectory() as temp_dir:
                # Save and extract zip file
                temp_zip_path = os.path.join(temp_dir, uploaded_file.name)
                with open(temp_zip_path, "wb") as temp_file:
                    temp_file.write(uploaded_file.getbuffer())
                with zipfile.ZipFile(temp_zip_path, "r") as zip_ref:
                    zip_ref.extractall(temp_dir)
                
                # Collect all Excel file paths
                for root, _, files in os.walk(temp_dir):
                    for file_name in files:
                        if file_name.endswith(".xlsx"):
                            full_path = os.path.join(root, file_name)
                            extracted_files.append(full_path)
                            st.write(f"File extracted: {full_path}")  # Debugging message
    return extracted_files


def load_data(excel_files):
    """Loads data from the extracted Excel files."""
    all_data = pd.DataFrame()
    required_columns = ['ALLOCATED TO', 'STATUS', 'PRODUCT_DESCRIPTION', 'DATE', 'File Name']

    for file_path in excel_files:
        if os.path.exists(file_path):  # Validate file exists before processing
            try:
                workbook = load_workbook(file_path, data_only=True)
                sheet = workbook.active

                # Extract all data including hidden columns
                data = pd.DataFrame(sheet.values)

                # Set first row as header
                data.columns = data.iloc[0]
                data = data[1:]

                # Add 'File Name' column if missing
                if 'File Name' not in data.columns:
                    data['File Name'] = os.path.basename(file_path)

                # Ensure all required columns are present
                for col in required_columns:
                    if col not in data.columns:
                        data[col] = None

                # Append to combined DataFrame
                all_data = pd.concat([all_data, data[required_columns]], ignore_index=True)
            except Exception as e:
                st.error(f"Error processing file {file_path}: {e}")
        else:
            st.error(f"File not found: {file_path}")
    return all_data


def main():
    st.title("Live Data Tracking Dashboard")

    # File uploader for zip files
    uploaded_files = st.file_uploader("Choose zip files", type="zip", accept_multiple_files=True)

    # Button to load files
    if st.button("Load Files") and uploaded_files:
        # Extract and load Excel files from the zip archives
        excel_files = extract_files_from_zip(uploaded_files)
        if not excel_files:
            st.warning("No Excel files found in the uploaded zip files.")
            return

        all_data = load_data(excel_files)
        if all_data.empty:
            st.warning("No valid data found in the processed files.")
            return

        # Process and visualize the data
        st.subheader("Data Overview")
        st.dataframe(all_data)

        st.subheader("Summary Metrics")
        total_count = all_data['PRODUCT_DESCRIPTION'].count()
        completed_count = all_data[all_data['STATUS'] == 'COMPLETED']['STATUS'].count()
        pending_count = all_data[all_data['STATUS'] == 'PENDING']['STATUS'].count()

        st.write(f"Total Items: {total_count}")
        st.write(f"Completed Items: {completed_count}")
        st.write(f"Pending Items: {pending_count}")

        # Visualizations
        status_counts = all_data['STATUS'].value_counts()
        st.bar_chart(status_counts)

        st.subheader("Detailed Data by Allocated User")
        user_summary = all_data.groupby(['ALLOCATED TO', 'STATUS'])['PRODUCT_DESCRIPTION'].count().unstack(fill_value=0)
        st.table(user_summary)

    else:
        st.info("Please upload zip files to get started.")


if __name__ == "__main__":
    main()
