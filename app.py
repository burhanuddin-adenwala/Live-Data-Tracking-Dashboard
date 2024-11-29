import pandas as pd
import streamlit as st
from openpyxl import load_workbook
import zipfile

# Function to load data from multiple ZIP files
def load_data_from_multiple_zips(uploaded_zips):
    all_data = pd.DataFrame()
    required_columns = ['ALLOCATED TO', 'STATUS', 'PRODUCT_DESCRIPTION', 'DATE', 'File Name']

    for uploaded_zip in uploaded_zips:
        with zipfile.ZipFile(uploaded_zip) as z:
            for file_name in z.namelist():
                if file_name.endswith(".xlsx"):
                    with z.open(file_name) as file:
                        try:
                            workbook = load_workbook(file, data_only=True)
                            sheet = workbook.active
                            data = pd.DataFrame(sheet.values)
                            data.columns = data.iloc[0]
                            data = data[1:]

                            # Add 'File Name' column if missing
                            if 'File Name' not in data.columns:
                                data['File Name'] = file_name

                            # Ensure all required columns are present
                            for col in required_columns:
                                if col not in data.columns:
                                    data[col] = None

                            all_data = pd.concat([all_data, data[required_columns]], ignore_index=True)
                        except Exception as e:
                            st.warning(f"Error processing {file_name}: {e}")

    return all_data

# Main application
def main():
    st.title("Enhanced Live Data Tracking Dashboard")

    # File uploader for multiple ZIP files
    uploaded_zips = st.file_uploader("Choose one or more zip folders containing Excel files", type="zip", accept_multiple_files=True)

    if st.button("Load Files") and uploaded_zips:
        # Load data from ZIP files
        all_data = load_data_from_multiple_zips(uploaded_zips)

        if not all_data.empty:
            # Normalize 'STATUS' values to uppercase
            all_data['STATUS'] = all_data['STATUS'].str.upper()

            # Create a user summary
            st.subheader("Detailed Data Overview")
            user_summary = all_data.groupby(['File Name', 'ALLOCATED TO']).agg(
                Total_Count=('PRODUCT_DESCRIPTION', 'count'),
                Completed_Count=('STATUS', lambda x: (x == 'COMPLETED').sum()),
                Pending_Count=('STATUS', lambda x: (x == 'PENDING').sum())
            ).reset_index()

            # Create a pivot table for date-wise counts
            date_counts = all_data.pivot_table(index=['File Name', 'ALLOCATED TO'],
                                               columns='DATE',
                                               values='STATUS',
                                               aggfunc='size').fillna(0)

            # Merge date counts with user summary
            user_summary = user_summary.merge(date_counts, on=['File Name', 'ALLOCATED TO'], how='left')

            # Calculate "Actual Pending"
            date_sums = date_counts.sum(axis=1).reindex(
                user_summary.set_index(['File Name', 'ALLOCATED TO']).index, fill_value=0
            )
            user_summary['Actual Pending'] = user_summary['Pending_Count'] - date_sums.values

            # Add Grand Total row
            total_row = user_summary.select_dtypes(include='number').sum()
            total_row['File Name'] = 'Grand Total'
            total_row['ALLOCATED TO'] = '-'
            user_summary = pd.concat([user_summary, pd.DataFrame([total_row])], ignore_index=True)

            # Display the complete table
            st.dataframe(user_summary, use_container_width=True)
        else:
            st.warning("No data found in the uploaded files.")
    else:
        st.info("Please upload one or more zip folders to get started.")

if __name__ == "__main__":
    main()
