import pandas as pd
import streamlit as st
from ui import setup_sidebar, upload_excel_file
from file_handler import extract_order_info, compute_estimated_test_date, append_df_to_original_excel


st.set_page_config(page_title="订单信息提取", layout="wide")  # ✅ 最上方

def main():
    setup_sidebar()

    # ✅ 一开始就显示
    uploaded_file = upload_excel_file()

    if uploaded_file:
    if st.button("📥 生成订单信息"):
        df_info = extract_order_info(uploaded_file)
        df_info = compute_estimated_test_date(df_info)

        st.write("✅ 提取并计算结果：")
        st.dataframe(df_info)

        # ✅ 生成带原始数据的新 Excel 文件
        new_excel_bytes = append_df_to_original_excel(uploaded_file, df_info, new_sheet_name="提取结果")

        st.download_button(
            label="📥 下载含提取结果的完整 Excel",
            data=new_excel_bytes,
            file_name="提取结果_完整版本.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


if __name__ == "__main__":
    main()
