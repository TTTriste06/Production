import pandas as pd
import streamlit as st
from ui import setup_sidebar, upload_excel_file
from file_handler import extract_order_info, write_xyz_columns, compute_estimated_test_date, write_calendar_headers


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
    
            updated_file = write_xyz_columns(uploaded_file, df_info)
            updated_file = write_calendar_headers(updated_file, df_info)

    
            st.download_button(
                label="📥 下载更新后的原始 Excel",
                data=updated_file,
                file_name="更新后的封装表.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if __name__ == "__main__":
    main()
