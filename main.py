import streamlit as st
from ui import setup_sidebar, upload_excel_file
from file_handler import extract_target_fields_from_sheet1

def main():
    st.set_page_config(page_title="订单信息提取", layout="wide")
    setup_sidebar()

    uploaded_file = upload_excel_file()

    if uploaded_file:
        extracted_df = extract_target_fields_from_sheet1(uploaded_file)
        st.write("✅ 提取结果：")
        st.write(extracted_df)

if __name__ == "__main__":
    main()
