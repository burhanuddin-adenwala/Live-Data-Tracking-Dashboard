import pandas as pd
import streamlit as st
import zipfile
from openpyxl import load_workbook
import logging
import io  # Import io for byte stream handling

# Set up logging
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

# Function to process a single Excel file from a byte stream
def process_excel_file(file_bytes):
    try:
        # Load workbook from byte stream
        workbook = load_workbook(io.BytesIO(file_bytes), data_only=True)
        sheet = workbook.active
        data = pd.DataFrame(sheet.values)
        data.columns = data.iloc[0]  # Set first row as header
        data = data[1:]  # Drop header row
        return data
    except Exception as e:
        logger.error(f"Error processing file: {e}")
        st.warning(f"Error processing file: {e}")
        return None

# Main application
def main():
    st.title("Enhanced Live Data Tracking Dashboard")

    # File uploader for ZIP file
    uploaded_zip = st.file_uploader("Choose a zip folder containing Excel files", type="zip")

    if uploaded_zip:
        # Extract files from the ZIP
        with zipfile.ZipFile(uploaded_zip, 'r') as z:
            file_names = [f for f in z.namelist() if f.endswith(".xlsx")]

        if not file_names:
            st.warning("No Excel files found in the ZIP.")
            return

        # Initialize an empty DataFrame to collect all data
        all_data = pd.DataFrame()

        # Create a progress bar
        progress_bar = st.progress(0)
        progress_text = st.empty()

        # Process each file one by one
        for i, file_name in enumerate(file_names):
            with z.open(file_name) as file:
                progress_text.text(f"Processing file {i + 1} of {len(file_names)}: {file_name}")

                # Read the file bytes and process the Excel file
                file_bytes = file.read()
                data = process_excel_file(file_bytes)
                if data is not None:
                    all_data = pd.concat([all_data, data], ignore_index=True)

            # Update progress bar
            progress_bar.progress((i + 1) / len(file_names))

        # Show the processed data after all files are loaded
        if not all_data.empty:
            st.subheader("Processed Data")
            st.dataframe(all_data)

            # Data cleaning and transformation
            required_columns = ['ALLOCATED TO', 'STATUS', 'PRODUCT_DESCRIPTION', 'DATE', 'File Name']
            missing_columns = [col for col in required_columns if col not in all_data.columns]

            if missing_columns:
                st.error(f"Missing columns: {missing_columns}")
                return

            # Clean data
            all_data['STATUS'] = all_data['STATUS'].str.upper()  # Normalize 'STATUS'
            all_data['DATE'] = pd.to_datetime(all_data['DATE'], errors='coerce')  # Convert 'DATE' to datetime
            all_data.dropna(subset=['ALLOCATED TO', 'STATUS', 'DATE'], inplace=True)

            # Create user summary
            user_summary = all_data.groupby(['File Name', 'ALLOCATED TO']).agg(
                Total_Count=('PRODUCT_DESCRIPTION', 'count'),
                Completed_Count=('STATUS', lambda x: (x == 'COMPLETED').sum()),
                Pending_Count=('STATUS', lambda x: (x == 'PENDING').sum())
            ).reset_index()

            # Display the summary table
            st.subheader("User Summary")
            st.dataframe(user_summary)
        else:
            st.warning("No valid data found after processing the files.")

if __name__ == "__main__":
    main()
