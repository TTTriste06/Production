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


def add_headers_to_xyz(excel_file: BytesIO) -> BytesIO:
    """
    在 X、Y、Z 列添加两行标题：
    第3行：X=预估开始测试日期，Y=结束日期，Z=日期
    第4行：X=预估开始测试日期，Y=结束日期，Z=星期
    """
    wb = load_workbook(excel_file)
    ws = wb["Sheet1"]  # 只处理 Sheet1，必要时可参数化

    # 第3、4行的 X, Y, Z 列是第 24, 25, 26 列
    columns = {
        24: ["预估开始测试日期", "预估开始测试日期"],
        25: ["结束日期", "结束日期"],
        26: ["日期", "星期"]
    }

    for col, values in columns.items():
        for i, val in enumerate(values):  # i=0 -> row=3, i=1 -> row=4
            ws.cell(row=3+i, column=col, value=val)

    # 保存回 BytesIO
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
    - '实际开始测试日'：日期列
    - '排产周期'、'磨划周期'、'封装周期'：整数字段，单位为天
    """
    # 复制 DataFrame 防止原地修改
    df = df.copy()

    # 确保日期格式正确
    df["实际开始测试日"] = pd.to_datetime(df["实际开始测试日"], errors="coerce")

    # 填充空周期为 0
    for col in ["排产周期", "磨划周期", "封装周期"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)

    # 计算总周期天数
    df["总周期天数"] = df["排产周期"] + df["磨划周期"] + df["封装周期"]

    # 计算预估开始测试日期
    df["预估开始测试日期"] = df["实际开始测试日"] + pd.to_timedelta(df["总周期天数"], unit="D")

    # 格式化为 yyyy/mm/dd 字符串
    df["预估开始测试日期"] = df["预估开始测试日期"].dt.strftime("%Y/%m/%d")

    return df

