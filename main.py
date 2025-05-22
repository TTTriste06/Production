import streamlit as st
from ui import setup_sidebar, upload_excel_file
from file_handler import extract_order_info, compute_estimated_test_date

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

            from io import BytesIO
            def to_excel(df):
                output = BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    df.to_excel(writer, index=False, sheet_name="订单信息")
                output.seek(0)
                return output

            excel_bytes = to_excel(df_info)
            st.download_button(
                label="📥 下载生成的 Excel 文件",
                data=excel_bytes,
                file_name="提取结果.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if __name__ == "__main__":
    main()
