import streamlit as st
import pandas as pd
from io import BytesIO
from scheduler import schedule_sheet

st.set_page_config(page_title="封装排产计划生成器", layout="wide")
st.title("📦 委外封装排产软件")

uploaded_file = st.file_uploader("上传订单 Excel 文件（包含排产字段）", type=["xlsx"])

if uploaded_file:
    df_raw = pd.read_excel(uploaded_file, sheet_name=0, header=None)
    st.success("✅ 文件上传成功！开始解析...")

    # 从第5行作为字段行，第6行开始是数据
    header_row = 4
    df_raw.columns = df_raw.iloc[header_row]
    df_data = df_raw.iloc[header_row+1:].copy()
    df_data.reset_index(drop=True, inplace=True)

    # 检查必要字段
    required_columns = ["产品", "订单数", "封装厂", "封装形式", "wafer in", "排产周期", "磨划周期", "封装周期", "分配产能"]
    missing = [col for col in required_columns if col not in df_data.columns]
    if missing:
        st.error(f"❌ 缺少必要字段：{missing}")
    else:
        # 调用排产逻辑
        try:
            df_scheduled = schedule_sheet(df_data)
            st.success("✅ 排产完成！")
            st.dataframe(df_scheduled.head())

            # 导出 Excel
            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                df_scheduled.to_excel(writer, sheet_name="排产计划", index=False)
            st.download_button("📥 下载排产结果", data=output.getvalue(), file_name="排产计划结果.xlsx")
        except ValueError as e:
            st.error(f"❌ 排产失败：{e}")
