import pandas as pd

TARGET_COLUMNS = [
    "订单号", "封装厂", "封装形式", "wafer in", "需排产",
    "排产周期", "磨划周期", "封装周期", "总产能", "分配产能", "实际开始测试日"
]

def extract_order_info(uploaded_file):
    """
    从上传的 Excel 中提取 Sheet1 的目标字段。
    
    参数:
        uploaded_file: Streamlit 上传的文件对象

    返回:
        pd.DataFrame: 包含目标字段的数据
    """
    try:
        df = pd.read_excel(uploaded_file, sheet_name="Sheet1", header=3)
        matching_columns = [col for col in TARGET_COLUMNS if col in df.columns]
        df_filtered = df[matching_columns].copy()
        return df_filtered
    except Exception as e:
        return pd.DataFrame({"错误信息": [str(e)]})


def compute_estimated_test_date(df: pd.DataFrame) -> pd.DataFrame:
    """
    根据 wafer in 和周期列计算预估开始测试日期，并添加新列。
    
    参数:
        df: 包含原始字段的 DataFrame（需包含 wafer in 和周期列）

    返回:
        pd.DataFrame: 增加了“预估开始测试日期”和“结束日期”的 DataFrame
    """
    try:
        df = df.copy()
        df["wafer in"] = pd.to_datetime(df["wafer in"], errors="coerce")
        df["总周期"] = df[["排产周期", "磨划周期", "封装周期"]].sum(axis=1, skipna=True)
        df["预估开始测试日期"] = df["wafer in"] + pd.to_timedelta(df["总周期"], unit="D")
        df["结束日期"] = df["预估开始测试日期"]  # 占位（你可修改逻辑）
        return df
    except Exception as e:
        return pd.DataFrame({"错误信息": [str(e)]})
