import streamlit as st
from ui import setup_sidebar, upload_excel_file
from file_handler import extract_order_info, compute_estimated_test_date

def main():
    st.set_page_config(page_title="订单信息提取", layout="wide")
    setup_sidebar()

    uploaded_file = upload_excel_file()

    if uploaded_file:
        generate = st.button("📥 生成订单信息")
        if generate:
            df_info = extract_order_info(uploaded_file)
            df_info = compute_estimated_test_date(df_info)
            st.write("✅ 提取并计算结果：")
            st.write(df_info)

if __name__ == "__main__":
    main()




