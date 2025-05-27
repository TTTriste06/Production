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
        

def adjust_column_width_for_openpyxl(ws, df, start_col=25):
    for i, col in enumerate(df.columns):
        col_letter = get_column_letter(start_col + i)
        content_max_len = (
            df[col].dropna().astype(str).str.len().max()
            if not df[col].dropna().empty
            else 0
        )
        header_len = len(str(col))
        width = min(max(content_max_len, header_len) * 1.2 + 7, 50)
        ws.column_dimensions[col_letter].width = width

def write_xyz_columns(excel_file: BytesIO, df_info: pd.DataFrame) -> BytesIO:
    """
    写入 X、Y、Z 列：
    - 第3/4行写入标题
    - X列填入 '预估开始测试日期'
    - Z列填入固定值 "排产"
    - 自动调整列宽
    """
    wb = load_workbook(excel_file)
    ws = wb["Sheet1"]

    # 表头写入（行3和4）
    header_map = {
        24: ["预估开始测试日期", "预估开始测试日期"],  # X
        25: ["结束日期", "结束日期"],          # Y（保留占位）
        26: ["日期", "星期"]                  # Z
    }

    for col_idx, values in header_map.items():
        for i, val in enumerate(values):
            ws.cell(row=3 + i, column=col_idx, value=val)

    # 写入 X列（第24列）内容
    for i, value in enumerate(df_info["预估开始测试日期"], start=5):
        ws.cell(row=i, column=24, value=value)

    # 写入 Z列（第26列）内容为固定值“排产”
    for i in range(5, 5 + len(df_info)):
        ws.cell(row=i, column=26, value="排产")

    # 调整列宽
    temp_df = pd.DataFrame({
        "预估开始测试日期": df_info["预估开始测试日期"],
        "排产": ["排产"] * len(df_info)
    })
    adjust_column_width_for_openpyxl(ws, temp_df, start_col=24)

    # 保存为 BytesIO 返回
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

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


def write_calendar_headers(excel_file: BytesIO, df_info: pd.DataFrame, start_col: int = 27, weeks: int = 7) -> BytesIO:
    """
    从指定起始列写入日期与星期，并自动调整这些列的列宽。
    - 第3行：日期（yyyy/mm/dd）
    - 第4行：星期（中文 一~日）
    """
    wb = load_workbook(excel_file)
    ws = wb["Sheet1"]

    df_info = df_info.copy()
    df_info["预估开始测试日期"] = pd.to_datetime(df_info["预估开始测试日期"], errors="coerce")
    min_date = df_info["预估开始测试日期"].min()
    if pd.isna(min_date):
        raise ValueError("❌ 无法从 '预估开始测试日期' 中解析出合法日期")

    weekday_map = {0: "一", 1: "二", 2: "三", 3: "四", 4: "五", 5: "六", 6: "日"}

    date_list = []
    weekday_list = []

    for i in range(weeks * 7):
        col_idx = start_col + i
        current_date = min_date + pd.Timedelta(days=i)

        formatted_date = current_date.strftime("%Y/%m/%d")
        formatted_weekday = weekday_map[current_date.weekday()]

        ws.cell(row=3, column=col_idx, value=formatted_date)
        ws.cell(row=4, column=col_idx, value=formatted_weekday)

        date_list.append(formatted_date)
        weekday_list.append(formatted_weekday)

    # 构建临时 DataFrame 用于列宽调整
        temp_df = pd.DataFrame({
        "日期": date_list,
        "星期": weekday_list
    })

    adjust_column_width_for_openpyxl(ws, temp_df.T, start_col=start_col)

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output
