import streamlit as st
import pandas as pd
import numpy as np # Not used in the example, but kept if you plan to use it later
import time # Not used in the example, but kept if you plan to use it later
import io

st.set_page_config(page_title="Uber Pickups Data App", layout="centered")
st.title('try the new')
st.write("Upload your CSV or XLSX file below to view its contents and download the processed version.")

# The file uploader
uploaded_file = st.file_uploader("Choose a file", type=["csv", "xlsx"])

# This entire block will ONLY run if a file has been uploaded
if uploaded_file is not None:
    try:
        # Read the file based on its type
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
            st.success("CSV file uploaded successfully!")
        elif uploaded_file.name.endswith('.xlsx'):
            df = pd.read_excel(uploaded_file)
            st.success("XLSX file uploaded successfully!")
        else:
            st.error("Unsupported file type. Please upload a CSV or XLSX file.")
            st.stop() # Stop execution if the file type is wrong


        # --- Data Processing Section ---
        st.markdown("---")
        st.write("### Processing Data...")

        # Your processing step: adding a new column '0' with value 29
        df['0'] = 29
        st.write("### Processed File Content (first 5 rows):")
        # st.dataframe(df.head()) # Show the processed head

        # --- Download Section ---
        st.markdown("---")
        st.write("### Download the Processed File:")

        excel_buffer = io.BytesIO()
        df.to_excel(excel_buffer, index=False, engine='openpyxl')
        excel_buffer.seek(0) # Rewind the buffer to the beginning

        st.download_button(
            label="Download Processed XLSX",
            data=excel_buffer,
            file_name="processed_uber_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_excel_button"
        )
        print(df.head()) # This print statement will now only execute after upload and processing

    except Exception as e:
        st.error(f"An error occurred while processing the file: {e}")
else:
    # Message to display when no file is uploaded yet
    st.info("Please upload a file to proceed with data processing and download.")