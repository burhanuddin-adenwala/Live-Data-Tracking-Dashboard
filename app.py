import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
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
            st.subheader("Detailed User Information by File Name")
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

            # Calculate the "Difference" column
            date_sums = date_counts.sum(axis=1).reindex(
                user_summary.set_index(['File Name', 'ALLOCATED TO']).index, fill_value=0
            )
            user_summary['Difference'] = user_summary['Completed_Count'] - date_sums.values

            # Add Grand Total row
            total_row = user_summary.select_dtypes(include='number').sum()
            total_row['File Name'] = 'Grand Total'
            total_row['ALLOCATED TO'] = '-'
            user_summary = pd.concat([user_summary, pd.DataFrame([total_row])], ignore_index=True)

            # Display the summary table
            st.dataframe(user_summary, use_container_width=True)

            # Status Overview: Bar chart for Completed vs Pending Counts
            st.subheader("Status Overview")
            status_counts = user_summary[['File Name', 'ALLOCATED TO', 'Completed_Count', 'Pending_Count']].dropna()
            status_counts.set_index(['File Name', 'ALLOCATED TO']).plot(kind='bar', stacked=True, figsize=(10, 6))
            plt.title('Completed vs Pending Counts')
            plt.ylabel('Count')
            plt.xlabel('File Name and User')
            st.pyplot(plt)

            # Status Distribution: Pie chart
            st.subheader("Status Distribution")
            status_distribution = all_data['STATUS'].value_counts()
            plt.figure(figsize=(6, 6))
            plt.pie(status_distribution, labels=status_distribution.index, autopct='%1.1f%%', startangle=90,
                    colors=['#4CAF50', '#FF9800'])
            plt.title("Distribution of Completed and Pending Tasks")
            st.pyplot(plt)

            # Total Product Count per User: Bar chart
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
        st.info("Please upload one or more zip folders to get started.")

if __name__ == "__main__":
    main()
