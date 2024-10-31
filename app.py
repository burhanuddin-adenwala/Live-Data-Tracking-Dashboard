import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt

def load_data(uploaded_files):
    # Initialize an empty DataFrame to store combined data
    all_data = pd.DataFrame()

    # Loop through uploaded files
    for uploaded_file in uploaded_files:
        if uploaded_file.name.endswith(".xlsx"):
            data = pd.read_excel(uploaded_file)

            # Add the filename to the data
            data['File Name'] = uploaded_file.name
            all_data = pd.concat([all_data, data[['ALLOCATED TO', 'STATUS', 'PRODUCT_DESCRIPTION', 'DATE', 'File Name']]], ignore_index=True)

    return all_data

def main():
    st.title("Live Data Tracking Dashboard")

    # File uploader for Excel files
    uploaded_files = st.file_uploader("Choose Excel files", type="xlsx", accept_multiple_files=True)

    # Button to load files
    if st.button("Load Files") and uploaded_files:
        all_data = load_data(uploaded_files)

        if not all_data.empty:
            # Ensure STATUS is capitalized
            all_data['STATUS'] = all_data['STATUS'].str.upper()

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
        st.info("Please upload Excel files to get started.")

if __name__ == "__main__":
    main()
