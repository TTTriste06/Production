import streamlit as st
import pandas as pd
from io import BytesIO
from file_handler import extract_order_info, compute_estimated_test_date

def to_excel(df: pd.DataFrame) -> BytesIO:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="订单信息")
    output.seek(0)
    return output

def main():
    st.set_page_config(page_title="订单信息提取", layout="wide")

    st.sidebar.title("📊 Excel 工具")
    st.sidebar.markdown("上传封装交付表 → 提取 → 生成 → 下载")

    uploaded_file = st.file_uploader("📤 上传 Excel 文件", type=["xlsx"])

    if uploaded_file:
        if st.button("📥 生成订单信息"):
            df_info = extract_order_info(uploaded_file)
            df_info = compute_estimated_test_date(df_info)

            st.write("✅ 提取并计算结果：")
            st.dataframe(df_info)

            # 导出为 Excel 并生成下载链接
            excel_bytes = to_excel(df_info)
            st.download_button(
                label="📥 下载生成的 Excel 文件",
                data=excel_bytes,
                file_name="提取结果.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if __name__ == "__main__":
    main()
