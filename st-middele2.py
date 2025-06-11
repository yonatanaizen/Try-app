import streamlit as st
import pandas as pd
import io
from process_the_data import   all_t


def data_process(uploaded_file:pd.DataFrame):
    '''
    funcion that got data and do the process
    :param data: dataframe
    :return: df new column 9
    '''
    if uploaded_file.name.endswith('.csv'):
        data=pd.read_csv(uploaded_file)
    else:
        data=pd.read_excel(uploaded_file)

    data['new_c']=9
    return data

st.set_page_config(page_title="Montly result", layout="centered")
st.title('Monthly result')
st.write("Upload your CSV or XLSX file")

uploaded_file = st.file_uploader("Choose a file", type=["csv", "xlsx"])

# This entire block will ONLY run if a file has been uploaded
if uploaded_file is not None:
    wb = all_t(uploaded_file)  # wb is an openpyxl.Workbook

    st.markdown("---")
    st.write("### Download the Processed File:")

    excel_buffer = io.BytesIO()
    wb.save(excel_buffer)
    excel_buffer.seek(0)

    st.download_button(
        label="Download Processed XLSX",
        data=excel_buffer,
        file_name="data_after_process.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download_excel_button"
    )