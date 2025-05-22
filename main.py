import streamlit as st
from ui import setup_sidebar, upload_excel_file
from file_handler import read_excel_file

def main():
    st.set_page_config(page_title="Excel 文件预览", layout="wide")
    setup_sidebar()

    uploaded_file = upload_excel_file()

    if uploaded_file:
        excel_data = read_excel_file(uploaded_file)

        for sheet_name, df in excel_data.items():
            st.subheader(f"📄 工作表: {sheet_name}")
            st.dataframe(df)

if __name__ == "__main__":
    main()
