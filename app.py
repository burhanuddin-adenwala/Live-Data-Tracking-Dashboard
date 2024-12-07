import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
from openpyxl import load_workbook
import zipfile
import tempfile

def extract_files_from_zip(uploaded_files):
    excel_files = []
    for uploaded_file in uploaded_files:
        if uploaded_file.name.endswith(".zip"):
            with tempfile.TemporaryDirectory() as temp_dir:
                # Save and extract zip file
                with open(f"{temp_dir}/{uploaded_file.name}", "wb") as temp_file:
                    temp_file.write(uploaded_file.getbuffer())
                with zipfile.ZipFile(f"{temp_dir}/{uploaded_file.name}", "r") as zip_ref:
                    zip_ref.extractall(temp_dir)
                
                # Collect all Excel files
                for file in zip_ref.namelist():
                    if file.endswith(".xlsx"):
                        excel_files.append(f"{temp_dir}/{file}")
    return excel_files

def load_data(excel_files):
    all_data = pd.DataFrame()
    required_columns = ['ALLOCATED TO', 'STATUS', 'PRODUCT_DESCRIPTION', 'DATE', 'File Name']

    for file_path in excel_files:
        # Load workbook with openpyxl and include hidden columns
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

    return all_data

def main():
    st.title("Live Data Tracking Dashboard")

    # File uploader for zip folders
    uploaded_files = st.file_uploader("Upload zip folders containing Excel files", type="zip", accept_multiple_files=True)

    if st.button("Load Files") and uploaded_files:
        # Extract Excel files from zip folders
        excel_files = extract_files_from_zip(uploaded_files)

        if excel_files:
            all_data = load_data(excel_files)

            if not all_data.empty:
                # Ensure STATUS is capitalized
                all_data['STATUS'] = all_data['STATUS'].str.upper()

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

                # Add "Difference" and "Actual Pending Count" columns
                user_summary['Date-wise Sum'] = date_counts.sum(axis=1).values
                user_summary['Difference'] = user_summary['Completed_Count'] - user_summary['Date-wise Sum']
                user_summary['Actual Pending Count'] = user_summary['Total_Count'] - user_summary['Date-wise Sum']

                # Display the summary table
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
                st.warning("No data found in the extracted Excel files.")
        else:
            st.warning("No Excel files found in the uploaded zip folders.")
    else:
        st.info("Please upload zip folders to get started.")

if __name__ == "__main__":
    main()
