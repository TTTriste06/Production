import pandas as pd

TARGET_COLUMNS = [
    "订单号", "封装厂", "封装形式", "waferin", "需排产",
    "排产周期", "磨划周期", "封装周期", "总产能", "分配产能", "实际开始测试日期"
]

def extract_target_fields_from_sheet1(uploaded_file):
    try:
        # 只读 Sheet1，并指定 header 行
        df = pd.read_excel(uploaded_file, sheet_name="Sheet1", header=3)

        # 提取目标列
        matching_columns = [col for col in TARGET_COLUMNS if col in df.columns]
        if not matching_columns:
            return pd.DataFrame({"提示": ["未在 Sheet1 中找到指定字段"]})
        
        return df[matching_columns].dropna(how="all")  # 去除全为空的行

    except Exception as e:
        return pd.DataFrame({"错误信息": [str(e)]})
