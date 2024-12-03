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
        # Open the workbook and unhide columns
        workbook = load_workbook(file, data_only=True)
        sheet = workbook.active

        # Unhide all columns
        for col_dim in sheet.column_dimensions.values():
            col_dim.hidden = False

        # Remove filters if applied
        sheet.auto_filter.ref = None

        # Convert the worksheet to a DataFrame
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
                            # Load the Excel file and process data
                            data = load_excel_file(file)

                            # Add 'File Name' column if missing
                            if 'File Name' not in data.columns:
                                data['File Name'] = file_name

                            # Ensure all required columns are present
                            for col in required_columns:
                                if col not in data.columns:
                                    data[col] = None

                            # Concatenate to the main DataFrame
                            all_data = pd.concat([all_data, data[required_columns]], ignore_index=True)
                    except Exception as e:
                        logger.error(f"Error processing file {file_name}: {e}")
                        st.warning(f"Skipping file {file_name} due to an error: {e}")
    return all_data

# Main application
def main():
    st.title("Enhanced Live Data Tracking Dashboard")

    # File uploader for multiple ZIP files
    uploaded_zips = st.file_uploader(
        "Choose one or more zip folders containing Excel files",
        type="zip",
        accept_multiple_files=True
    )

    if st.button("Load Files") and uploaded_zips:
        # Load data from ZIP files
        all_data = load_data_from_multiple_zips(uploaded_zips)

        if not all_data.empty:
            # Check required columns
            required_columns = ['ALLOCATED TO', 'STATUS', 'PRODUCT_DESCRIPTION', 'DATE', 'File Name']
            missing_columns = [col for col in required_columns if col not in all_data.columns]
            if missing_columns:
                st.error(f"Missing columns in data: {missing_columns}")
                return

            # Clean data
            all_data['STATUS'] = all_data['STATUS'].str.upper()  # Normalize 'STATUS' column
            all_data['DATE'] = pd.to_datetime(all_data['DATE'], errors='coerce')  # Ensure 'DATE' is datetime
            all_data.dropna(subset=['ALLOCATED TO', 'STATUS', 'DATE'], inplace=True)  # Drop rows with critical NaNs

            if all_data.empty:
                st.warning("No valid data found after cleaning.")
                return

            try:
                # Create a pivot table for date-wise counts
                date_counts = all_data.pivot_table(
                    index=['File Name', 'ALLOCATED TO'],
                    columns='DATE',
                    values='STATUS',
                    aggfunc='size'
                ).fillna(0)

            except Exception as e:
                logger.error(f"Error creating pivot table: {e}")
                st.error(f"Error creating pivot table: {e}")
                st.text(f"Data preview for debugging:\n{all_data.head()}")
                return

            # Create a user summary
            user_summary = all_data.groupby(['File Name', 'ALLOCATED TO']).agg(
                Total_Count=('PRODUCT_DESCRIPTION', 'count'),
                Completed_Count=('STATUS', lambda x: (x == 'COMPLETED').sum()),
                Pending_Count=('STATUS', lambda x: (x == 'PENDING').sum())
            ).reset_index()

            # Merge date counts with user summary
            user_summary = user_summary.merge(date_counts, on=['File Name', 'ALLOCATED TO'], how='left')

            # Calculate additional columns
            date_sums = date_counts.sum(axis=1).reindex(
                user_summary.set_index(['File Name', 'ALLOCATED TO']).index, fill_value=0
            )
            user_summary['Difference'] = user_summary['Completed_Count'] - date_sums.values
            user_summary['Actual Pending'] = user_summary['Total_Count'] - date_sums.values

            # Add Grand Total row
            total_row = user_summary.select_dtypes(include='number').sum()
            total_row['File Name'] = 'Grand Total'
            total_row['ALLOCATED TO'] = '-'
            user_summary = pd.concat([user_summary, pd.DataFrame([total_row])], ignore_index=True)

            # Display the summary table
            st.subheader("Detailed User Information by File Name")
            st.dataframe(user_summary, use_container_width=True)
        else:
            st.warning("No data found in the uploaded files.")
    else:
        st.info("Please upload one or more zip folders to get started.")

if __name__ == "__main__":
    main()
