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
            # 1. 提取字段
            df_info = extract_order_info(uploaded_file)

            # 2. 计算预估开始测试日期
            df_info = compute_estimated_test_date(df_info)

            # 3. 生成 Excel 表头（X,Y,Z）
            updated_file = write_xyz_columns(uploaded_file, df_info)

            # 4. 写入日期/星期列（从 AB 开始），返回 updated_file 并生成 date_columns
            updated_file = write_calendar_headers(updated_file, df_info)

            # 5. 找到所有日期列名（从 AB 开始）作为排产目标列
            date_columns = [col for col in df_info.columns if col.startswith("20")]

            # 6. 排产逻辑处理：按封装厂+封装形式+产能安排每日产量
            # df_info = schedule_production(df_info, date_columns)

            # 7. 显示最终排产表（带日期列）
            # st.write("✅ 排产计划预览：")
            # st.dataframe(df_info)

            # 8. ✅ （可选）将含排产量的 df_info 再写入 Excel（此处你可以补一个函数）
            # updated_file = write_production_to_excel(updated_file, df_info, start_col=28)

            st.download_button(
                label="📥 下载排产计划 Excel",
                data=updated_file,
                file_name="排产结果.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if __name__ == "__main__":
    main()
