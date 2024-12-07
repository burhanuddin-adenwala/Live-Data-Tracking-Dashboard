import pandas as pd
import streamlit as st
import zipfile
from openpyxl import load_workbook
import logging

# Set up logging
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

# Function to load data from a single Excel file
def load_excel_file(file):
    try:
        workbook = load_workbook(file, data_only=True)
        sheet = workbook.active

        # Unhide all columns
        for col_dim in sheet.column_dimensions.values():
            col_dim.hidden = False

        # Remove filters if applied
        sheet.auto_filter.ref = None

        # Convert worksheet to DataFrame
        data = pd.DataFrame(sheet.values)
        data.columns = data.iloc[0]  # Set first row as header
        data = data[1:]  # Drop header row
        return data
    except Exception as e:
        logger.error(f"Error loading Excel file: {e}")
        raise

# Function to load data from multiple ZIP files
def load_data_from_multiple_zips(uploaded_zips):
    all_data = pd.DataFrame()
    required_columns = ['ALLOCATED TO', 'STATUS', 'PRODUCT_DESCRIPTION', 'DATE', 'File Name']

    for uploaded_zip in uploaded_zips:
        with zipfile.ZipFile(uploaded_zip) as z:
            for file_name in z.namelist():
                if file_name.endswith(".xlsx"):
                    try:
                        with z.open(file_name) as file:
                            data = load_excel_file(file)
                            if 'File Name' not in data.columns:
                                data['File Name'] = file_name
                            for col in required_columns:
                                if col not in data.columns:
                                    data[col] = None
                            all_data = pd.concat([all_data, data[required_columns]], ignore_index=True)
                    except Exception as e:
                        logger.error(f"Error processing file {file_name}: {e}")
                        st.warning(f"Skipping file {file_name} due to an error: {e}")
    return all_data

# Main application
def main():
    st.title("Enhanced EU Data Tracking Dashboard")

    uploaded_zips = st.file_uploader(
        "Choose one or more zip folders containing Excel files",
        type="zip",
        accept_multiple_files=True
    )

    if st.button("Load Files") and uploaded_zips:
        with st.spinner("Processing files..."):
            all_data = load_data_from_multiple_zips(uploaded_zips)

        if not all_data.empty:
            required_columns = ['ALLOCATED TO', 'STATUS', 'PRODUCT_DESCRIPTION', 'DATE', 'File Name']
            missing_columns = [col for col in required_columns if col not in all_data.columns]
            if missing_columns:
                st.error(f"Missing columns in data: {missing_columns}")
                return

            # Normalize columns
            all_data['STATUS'] = all_data['STATUS'].str.upper()
            all_data['DATE'] = pd.to_datetime(all_data['DATE'], errors='coerce')

            if all_data.empty:
                st.warning("No valid data found after cleaning.")
                return

            # Calculate totals, completed, and pending counts for each user and file
            user_summary = all_data.groupby(['File Name', 'ALLOCATED TO']).apply(
                lambda group: pd.Series({
                    'Total_Count': len(group),  # Total rows in the file
                    'Completed_Count': (group['STATUS'] == 'COMPLETED').sum(),  # Rows marked as COMPLETED
                    'Pending_Count': (group['STATUS'] == 'PENDING').sum()  # Rows marked as PENDING
                })
            ).reset_index()

            # Track date-wise completed counts
            date_counts = all_data.groupby(['File Name', 'ALLOCATED TO', 'DATE']).apply(
                lambda group: (group['STATUS'] == 'COMPLETED').sum()
            ).unstack(fill_value=0)

            # Merge summaries
            user_summary = user_summary.merge(date_counts, on=['File Name', 'ALLOCATED TO'], how='left')

            # Add Difference and Actual Pending columns
            user_summary['Difference'] = user_summary['Completed_Count'] - user_summary.drop(
                columns=['File Name', 'ALLOCATED TO', 'Total_Count', 'Completed_Count', 'Pending_Count']
            ).sum(axis=1)

            # Actual Pending count: Total count - sum of date counts (completed)
            user_summary['Actual Pending'] = user_summary['Total_Count'] - user_summary.drop(
                columns=['File Name', 'ALLOCATED TO', 'Total_Count', 'Completed_Count', 'Pending_Count', 'Difference']
            ).sum(axis=1)

            # Add Grand Total row
            total_row = user_summary.select_dtypes(include='number').sum()
            total_row['File Name'] = 'Grand Total'
            total_row['ALLOCATED TO'] = '-'
            user_summary = pd.concat([user_summary, pd.DataFrame([total_row])], ignore_index=True)

            st.subheader("Detailed User Information by File Name")
            st.dataframe(user_summary, use_container_width=True)

            # Visualization
            st.subheader("Completion Status Overview")
            status_counts = all_data['STATUS'].value_counts()
            st.bar_chart(status_counts)

            st.subheader("Date-wise Completion Overview")
            st.line_chart(date_counts.T)

            st.success("Data processed successfully!")
        else:
            st.warning("No data found in the uploaded files.")
    else:
        st.info("Please upload one or more zip folders to get started.")

if __name__ == "__main__":
    main()
