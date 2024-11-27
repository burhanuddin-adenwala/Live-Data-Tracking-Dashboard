import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
from openpyxl import load_workbook
import zipfile
import gc

# Function to process a single file and return summarized data
def process_file(file, file_name):
    try:
        workbook = load_workbook(file, data_only=True)
        sheet = workbook.active
        data = pd.DataFrame(sheet.values)
        data.columns = data.iloc[0]  # Use the first row as column names
        data = data[1:]  # Skip the first row

        # Ensure required columns exist
        required_columns = ['ALLOCATED TO', 'STATUS', 'PRODUCT_DESCRIPTION', 'DATE', 'File Name']
        for col in required_columns:
            if col not in data.columns:
                data[col] = None

        # Add 'File Name' column
        data['File Name'] = file_name

        # Summarize data for this file
        summary = data.groupby('ALLOCATED TO').agg(
            Total_Count=('PRODUCT_DESCRIPTION', 'count'),
            Completed_Count=('STATUS', lambda x: (x.str.upper() == 'COMPLETED').sum()),
            Pending_Count=('STATUS', lambda x: (x.str.upper() == 'PENDING').sum())
        ).reset_index()

        summary['File Name'] = file_name
        return summary

    except Exception as e:
        st.warning(f"Error processing {file_name}: {e}")
        return pd.DataFrame()  # Return empty DataFrame on failure

# Function to load and process data from multiple ZIP files
def load_data_from_multiple_zips(uploaded_zips):
    all_summaries = []  # Store summaries instead of raw data

    for uploaded_zip in uploaded_zips:
        with zipfile.ZipFile(uploaded_zip) as z:
            for file_name in z.namelist():
                if file_name.endswith('.xlsx'):
                    with z.open(file_name) as file:
                        # Process each file individually and append its summary
                        summary = process_file(file, file_name)
                        if not summary.empty:
                            all_summaries.append(summary)

                        # Free up memory
                        gc.collect()

    # Combine all summaries into a single DataFrame
    return pd.concat(all_summaries, ignore_index=True) if all_summaries else pd.DataFrame()

# Main application
def main():
    st.title("Enhanced Live Data Tracking Dashboard")

    # File uploader for multiple ZIP files
    uploaded_zips = st.file_uploader("Choose one or more zip folders containing Excel files", type="zip", accept_multiple_files=True)

    if st.button("Load Files") and uploaded_zips:
        # Load and process data from ZIP files
        user_summary = load_data_from_multiple_zips(uploaded_zips)

        if not user_summary.empty:
            # Add Grand Total row
            total_row = user_summary.select_dtypes(include='number').sum()
            total_row['ALLOCATED TO'] = 'Grand Total'
            total_row['File Name'] = '-'
            user_summary = pd.concat([user_summary, pd.DataFrame([total_row])], ignore_index=True)

            # Display the summary table
            st.subheader("Detailed User Information by File Name")
            st.dataframe(user_summary, use_container_width=True)

            # Status Overview: Bar chart for Completed vs Pending Counts
            st.subheader("Status Overview")
            status_counts = user_summary.groupby('ALLOCATED TO')[['Completed_Count', 'Pending_Count']].sum()
            status_counts.plot(kind='bar', stacked=True, figsize=(10, 6))
            plt.title('Completed vs Pending Counts')
            plt.ylabel('Count')
            plt.xlabel('Allocated To')
            st.pyplot(plt)

            # Total Product Count per User: Bar chart
            st.subheader("Total Product Count per User")
            product_counts = user_summary.groupby('ALLOCATED TO')['Total_Count'].sum()
            product_counts.plot(kind='bar', color='#2196F3', figsize=(10, 6))
            plt.title("Total Product Count per User")
            plt.ylabel("Total Count")
            plt.xlabel("Allocated To")
            st.pyplot(plt)

        else:
            st.warning("No valid data found in the uploaded files.")
    else:
        st.info("Please upload one or more zip folders to get started.")

if __name__ == "__main__":
    main()
