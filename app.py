import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
from openpyxl import load_workbook
import zipfile
import tempfile


def extract_files_from_zip(uploaded_files):
    """Extracts Excel files from uploaded zip files."""
    excel_files = []
    for uploaded_file in uploaded_files:
        if uploaded_file.name.endswith(".zip"):
            with tempfile.TemporaryDirectory() as temp_dir:
                # Save and extract zip file
                temp_zip_path = f"{temp_dir}/{uploaded_file.name}"
                with open(temp_zip_path, "wb") as temp_file:
                    temp_file.write(uploaded_file.getbuffer())
                with zipfile.ZipFile(temp_zip_path, "r") as zip_ref:
                    zip_ref.extractall(temp_dir)
                
                # Collect all Excel file paths
                for file_name in zip_ref.namelist():
                    if file_name.endswith(".xlsx"):
                        excel_files.append(f"{temp_dir}/{file_name}")
    return excel_files


def load_data(excel_files):
    """Loads data from the extracted Excel files and processes hidden columns."""
    all_data = pd.DataFrame()
    required_columns = ['ALLOCATED TO', 'STATUS', 'PRODUCT_DESCRIPTION', 'DATE', 'File Name']

    for file_path in excel_files:
        try:
            # Load workbook with openpyxl
            workbook = load_workbook(file_path, data_only=True)
            sheet = workbook.active

            # Extract all data, including hidden columns
            data = pd.DataFrame(sheet.values)

            # Set first row as header
            data.columns = data.iloc[0]
            data = data[1:]

            # If 'File Name' column is missing, add it
            if 'File Name' not in data.columns:
                data['File Name'] = file_path.split('/')[-1]

            # Ensure all required columns are present in the DataFrame
            for col in required_columns:
                if col not in data.columns:
                    data[col] = None  # Fill missing columns with None (or empty values)

            # Append data to the combined DataFrame
            all_data = pd.concat([all_data, data[required_columns]], ignore_index=True)
        except Exception as e:
            st.error(f"Error processing file {file_path}: {str(e)}")
            continue  # Skip this file and move to the next
    return all_data


def main():
    st.title("Live Data Tracking Dashboard")

    # File uploader for zip files
    uploaded_files = st.file_uploader("Choose zip files", type="zip", accept_multiple_files=True)

    # Button to load files
    if st.button("Load Files") and uploaded_files:
        # Extract and load Excel files from the zip archives
        excel_files = extract_files_from_zip(uploaded_files)
        all_data = load_data(excel_files)

        if not all_data.empty:
            # Ensure STATUS is capitalized
            all_data['STATUS'] = all_data['STATUS'].str.upper()

            # Add "Difference" column
            all_data['Difference'] = all_data.groupby(['File Name', 'ALLOCATED TO'])['STATUS'].transform(
                lambda x: (x == 'COMPLETED').sum()) - all_data.groupby(['File Name', 'ALLOCATED TO'])['DATE'].transform(
                'count'
            )

            # Add "Actual Pending" column
            all_data['Actual Pending'] = all_data.groupby(['File Name', 'ALLOCATED TO'])['PRODUCT_DESCRIPTION'].transform(
                'count') - all_data.groupby(['File Name', 'ALLOCATED TO'])['DATE'].transform('count')

            # Detailed User Information by File Name
            st.subheader("Detailed User Information by File Name")

            # Aggregate counts by File Name and User
            user_summary = all_data.groupby(['File Name', 'ALLOCATED TO']).agg(
                Total_Count=('PRODUCT_DESCRIPTION', 'count'),
                Completed_Count=('STATUS', lambda x: (x == 'COMPLETED').sum()),
                Pending_Count=('STATUS', lambda x: (x == 'PENDING').sum())
            ).reset_index()

            # Date-wise counts in columns
            date_counts = all_data.pivot_table(index=['File Name', 'ALLOCATED TO'], 
                                               columns='DATE', 
                                               values='STATUS', 
                                               aggfunc=lambda x: len(x)).fillna(0)

            # Merge the two DataFrames
            user_summary = user_summary.merge(date_counts, on=['File Name', 'ALLOCATED TO'], how='left')

            # Add "Difference" and "Actual Pending" columns
            user_summary['Difference'] = user_summary['Completed_Count'] - user_summary.iloc[:, 4:].sum(axis=1)
            user_summary['Actual Pending'] = user_summary['Total_Count'] - user_summary.iloc[:, 4:].sum(axis=1)

            # Show user summary with expanded width
            st.dataframe(user_summary, use_container_width=True)

            # Visualization of Completed and Pending Counts
            st.subheader("Status Overview")
            status_counts = user_summary[['File Name', 'ALLOCATED TO', 'Completed_Count', 'Pending_Count']]
            status_counts.set_index(['File Name', 'ALLOCATED TO']).plot(kind='bar', stacked=True, figsize=(10, 6))
            plt.title('Completed vs Pending Counts')
            plt.ylabel('Count')
            plt.xlabel('File Name and User')
            st.pyplot(plt)

            # Pie Chart for Status Distribution
            st.subheader("Status Distribution")
            status_distribution = all_data['STATUS'].value_counts()
            plt.figure(figsize=(6, 6))
            plt.pie(status_distribution, labels=status_distribution.index, autopct='%1.1f%%', startangle=90, colors=['#4CAF50', '#FF9800'])
            plt.title("Distribution of Completed and Pending Tasks")
            st.pyplot(plt)

            # Bar Chart for Total Product Count per User
            st.subheader("Total Product Count per User")
            product_counts = all_data.groupby('ALLOCATED TO')['PRODUCT_DESCRIPTION'].count().reset_index()
            plt.figure(figsize=(10, 6))
            plt.bar(product_counts['ALLOCATED TO'], product_counts['PRODUCT_DESCRIPTION'], color='#2196F3')
            plt.title("Total Product Count per User")
            plt.ylabel("Total Count")
            plt.xticks(rotation=45)
            st.pyplot(plt)

        else:
            st.warning("No data found in the uploaded files.")
    else:
        st.info("Please upload zip files to get started.")


if __name__ == "__main__":
    main()
