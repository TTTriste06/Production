import pandas as pd

TARGET_COLUMNS = [
    "订单号", "封装厂", "封装形式", "wafer in", "需排产",
    "排产周期", "磨划周期", "封装周期", "总产能", "分配产能", "实际开始测试日"
]

def read_excel_file(uploaded_file):
    try:
        excel_data = pd.read_excel(uploaded_file, sheet_name=None)  # 读取所有工作表
        selected_data = {}

        for sheet_name, df in excel_data.items():
            # 查找是否包含目标列中任意列
            matching_columns = [col for col in TARGET_COLUMNS if col in df.columns]
            if matching_columns:
                selected_df = df[matching_columns].copy()
                selected_data[sheet_name] = selected_df

        return selected_data if selected_data else {"提示": pd.DataFrame({"信息": ["未在文件中找到目标字段"]})}

    except Exception as e:
        return {"错误": pd.DataFrame({"错误信息": [str(e)]})}
