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

def append_df_to_original_excel(original_file, new_df, new_sheet_name="提取结果") -> BytesIO:
    """
    将新数据添加为原 Excel 文件中的一个新工作表，并返回内存中的 BytesIO 对象。
    
    参数:
    - original_file: Streamlit 上传的文件对象
    - new_df: 需追加写入的新 DataFrame
    - new_sheet_name: 新工作表名
    
    返回:
    - BytesIO: 包含原始内容 + 新工作表 的 Excel 文件对象
    """
    # 读入原始 Excel 的全部内容
    original_excel = pd.ExcelFile(original_file)
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl", mode="w") as writer:
        # 将原有 sheet 写入
        for sheet_name in original_excel.sheet_names:
            df_sheet = original_excel.parse(sheet_name)
            df_sheet.to_excel(writer, sheet_name=sheet_name, index=False)

        # 写入新提取 sheet
        new_df.to_excel(writer, sheet_name=new_sheet_name, index=False)

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

def update_sheet_preserving_styles(uploaded_file, df_with_estimates, start_col=25):
    from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
    import pandas as pd
    from openpyxl.utils import get_column_letter
    from io import BytesIO

    # 样式定义
    blue_fill = PatternFill(start_color='BDD7EE', end_color='BDD7EE', fill_type='solid')
    thin_border = Border(
        left=Side(style='thin', color='000000'),
        right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'),
        bottom=Side(style='thin', color='000000')
    )
    bold_font = Font(bold=True)
    center_align = Alignment(horizontal="center", vertical="center")
    date_number_format = 'yyyy/mm/dd'

    wb = load_workbook(uploaded_file)
    ws = wb["Sheet1"]

    new_columns = ["预估开始测试日期", "结束日期"]

    # 写入表头（第4行为空，第5行为列名）
    for idx, col_name in enumerate(new_columns):
        col_idx = start_col + idx

        # 第4行留空 + 黑边
        cell1 = ws.cell(row=4, column=col_idx)
        cell1.border = thin_border

        # 第5行列名 + 蓝底 + 黑边 + 居中加粗
        cell2 = ws.cell(row=5, column=col_idx, value=col_name)
        cell2.fill = blue_fill
        cell2.border = thin_border
        cell2.font = bold_font
        cell2.alignment = center_align

    # 写入数据（第6行起）
    for row_idx, row in enumerate(df_with_estimates.itertuples(index=False), start=6):
        for offset, col_name in enumerate(new_columns):
            value = getattr(row, col_name, "")
            col_idx = start_col + offset
            cell = ws.cell(row=row_idx, column=col_idx, value=value)

            # 设置统一样式
            cell.border = thin_border
            cell.alignment = center_align
            if isinstance(value, pd.Timestamp):
                cell.number_format = date_number_format

    # 自动列宽
    adjust_column_width_for_openpyxl(ws, df_with_estimates[new_columns], start_col=start_col)

    # 保存
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output
