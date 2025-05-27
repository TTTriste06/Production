import pandas as pd
from io import BytesIO
import pandas as pd
import streamlit as st
import copy
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import PatternFill, Border, Font, Alignment
from openpyxl.utils import get_column_letter


TARGET_COLUMNS = [
    "订单号", "封装厂", "封装形式", "waferin", "需排产",
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

def write_xyz_columns(excel_file: BytesIO, df_info: pd.DataFrame) -> BytesIO:
    """
    将 df_info 中的“预估开始测试日期”等数据写入原始 Excel 文件中的 X、Y、Z 列，并设置第3/4行的列标题。
    """
    wb = load_workbook(excel_file)
    ws = wb["Sheet1"]

    # 1. 写入第3、4行标题
    headers = {
        24: ["预估开始测试日期", "预估开始测试日期"],  # X列
        25: ["结束日期", "结束日期"],          # Y列（你可以后续添加）
        26: ["日期", "星期"]                  # Z列（你可以后续添加）
    }
    for col_idx, values in headers.items():
        for i, val in enumerate(values):
            ws.cell(row=3+i, column=col_idx, value=val)

    # 2. 写入数据：从第5行开始
    for i, value in enumerate(df_info["预估开始测试日期"], start=5):
        ws.cell(row=i, column=24, value=value)  # X列

    # 保存为 BytesIO
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output



def adjust_column_width_for_openpyxl(ws, df, start_col=25):
    """
    根据 DataFrame 的新列内容调整 openpyxl 工作表中指定起始列之后的列宽。

    参数:
    - ws: openpyxl 的 Worksheet 对象
    - df: 包含需写入新列的 DataFrame（包含列名和内容）
    - start_col: 起始列位置（默认从 AB=28 开始）
    """
    for i, col in enumerate(df.columns):
        col_letter = get_column_letter(start_col + i)
        content_max_len = df[col].astype(str).str.len().max()
        header_len = len(str(col))
        width = min(max(content_max_len, header_len) * 1.2 + 10, 50)
        ws.column_dimensions[col_letter].width = width

def compute_estimated_test_date(df):
    """
    计算“预估开始测试日期”：waferin 日期 + 排产周期 + 磨划周期 + 封装周期，单位为天。

    要求 df 中存在：
    - 'waferin'：日期列
    - '排产周期'、'磨划周期'、'封装周期'：整数字段，单位为天
    """
    # 复制 DataFrame 防止原地修改
    df = df.copy()

    # 确保日期格式正确
    df["waferin"] = pd.to_datetime(df["waferin"], errors="coerce")

    # 填充空周期为 0
    for col in ["排产周期", "磨划周期", "封装周期"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)

    # 计算总周期天数
    df["总周期天数"] = df["排产周期"] + df["磨划周期"] + df["封装周期"]

    # 计算预估开始测试日期
    df["预估开始测试日期"] = df["waferin"] + pd.to_timedelta(df["总周期天数"], unit="D")

    # 格式化为 yyyy/mm/dd 字符串
    df["预估开始测试日期"] = df["预估开始测试日期"].dt.strftime("%Y/%m/%d")

    return df


def write_calendar_headers(excel_file: BytesIO, df_info: pd.DataFrame, start_col: int = 28, weeks: int = 7) -> BytesIO:
    """
    从 AB 列（第28列）开始，写入：
    - 第3行：从预估开始测试日期最早日期起往后连续 dates（默认7周=49天）
    - 第4行：对应星期（中文：一二三四五六日）

    参数:
    - excel_file: BytesIO 原始文件
    - df_info: 包含 '预估开始测试日期' 列的 DataFrame
    - start_col: 起始列（默认AB=28）
    - weeks: 写入的周数（默认7周）
    """
    wb = load_workbook(excel_file)
    ws = wb["Sheet1"]

    # 解析最早日期
    df_info = df_info.copy()
    df_info["预估开始测试日期"] = pd.to_datetime(df_info["预估开始测试日期"], errors="coerce")
    min_date = df_info["预估开始测试日期"].min()

    if pd.isna(min_date):
        raise ValueError("❌ 无法从 '预估开始测试日期' 中解析出合法日期")

    # 星期中文映射
    weekday_map = {
        0: "一", 1: "二", 2: "三", 3: "四", 4: "五", 5: "六", 6: "日"
    }

    for i in range(weeks * 7):
        col_idx = start_col + i
        current_date = min_date + pd.Timedelta(days=i)

        # 写入第3行：日期
        ws.cell(row=3, column=col_idx, value=current_date.strftime("%Y/%m/%d"))

        # 写入第4行：星期
        weekday = weekday_map[current_date.weekday()]
        ws.cell(row=4, column=col_idx, value=weekday)

    # 保存并返回 BytesIO
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output


