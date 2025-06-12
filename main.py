import pandas as pd
import streamlit as st
from ui import setup_sidebar, upload_excel_file
from file_handler import (
    extract_order_info,
    compute_estimated_test_date,
    write_xyz_columns,
    write_calendar_headers
)
from schedule_production import schedule_production

st.set_page_config(page_title="订单信息排产工具", layout="wide")

def main():
    setup_sidebar()
    uploaded_file = upload_excel_file()

    if uploaded_file:
        if st.button("📥 生成排产计划"):
            file_bytes = uploaded_file.read()
            uploaded_file.seek(0)
        
            # 1. 提取字段
            df_info = extract_order_info(uploaded_file)

            # 2. 计算预估开始测试日期
            df_info = compute_estimated_test_date(df_info)

            # 3. 生成 XYZ 表头
            file_bytes = write_xyz_columns(file_bytes, df_info)

            # ✅ 4. 写入日期表头并更新 df_info
            file_bytes, df_info = write_calendar_headers(file_bytes, df_info)

            # ✅ 5. 提取新写入的日期列
            date_columns = [col for col in df_info.columns if col.startswith("20")]

            # ✅ 6. 排产
            df_info = schedule_production(df_info, date_columns)

            # ✅ 7. （可选）再次写入排产结果到 Excel
            # updated_file = write_production_to_excel(updated_file, df_info, start_col=28)

            # ✅ 8. 提供下载
            st.download_button(
                label="📥 下载排产计划 Excel",
                data=updated_file,
                file_name="排产结果.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if __name__ == "__main__":
    main()
