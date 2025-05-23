import pandas as pd
from io import BytesIO
import pandas as pd
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
        df["waferin"] = pd.to_datetime(df["waferin"], errors="coerce")
        df["总周期"] = df[["排产周期", "磨划周期", "封装周期"]].sum(axis=1, skipna=True)
        df["预估开始测试日期"] = df["waferin"] + pd.to_timedelta(df["总周期"], unit="D")
        df["结束日期"] = df["预估开始测试日期"]  # 占位（你可修改逻辑）
        return df
    except Exception as e:
        return pd.DataFrame({"错误信息": [str(e)]})

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

def copy_cell_style(src_cell, target_cell):
    try:
        if src_cell.has_style:
            if hasattr(src_cell, "font"): target_cell.font = src_cell.font
            if hasattr(src_cell, "border"): target_cell.border = src_cell.border
            if hasattr(src_cell, "fill"): target_cell.fill = src_cell.fill
            if hasattr(src_cell, "number_format"): target_cell.number_format = src_cell.number_format
            if hasattr(src_cell, "protection"): target_cell.protection = src_cell.protection
            if hasattr(src_cell, "alignment"): target_cell.alignment = src_cell.alignment
    except Exception as e:
        print(f"⚠️ 样式复制失败: {e}")  # 或者使用 logging.warning(...)


def update_sheet_preserving_styles(uploaded_file, df_with_estimates, start_col=28):  # start_col = AB = 28
    """
    在 Sheet1 中追加列并复制样式，保持格式一致。
    
    参数:
    - uploaded_file: 上传的 Excel 文件
    - df_with_estimates: 包含新列（预估开始测试日期、结束日期）的 DataFrame
    - start_col: 新字段起始列号（28 表示 AB 列）
    
    返回:
    - BytesIO: 新生成的 Excel 文件，包含原格式+追加字段
    """
    wb = load_workbook(uploaded_file)
    ws = wb["Sheet1"]

    # 新列标题
    new_columns = ["预估开始测试日期", "结束日期"]

    # =====================
    # ✅ 写入表头（第5行）
    # =====================
    header_row = 5
    for idx, col_name in enumerate(new_columns):
        col_idx = start_col + idx
        cell = ws.cell(row=header_row, column=col_idx, value=col_name)
        # 复制样式（参考前一列）
        ref_cell = ws.cell(row=header_row, column=col_idx - 1)
        copy_cell_style(ref_cell, cell)

    # =====================
    # ✅ 写入数据（从第6行开始）
    # =====================
    for row_idx, row in enumerate(df_with_estimates.itertuples(index=False), start=6):
        for col_offset, col_name in enumerate(new_columns):
            value = getattr(row, col_name, "")
            col_idx = start_col + col_offset
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            ref_cell = ws.cell(row=row_idx, column=col_idx - 1)
            copy_cell_style(ref_cell, cell)

    # 保存到内存
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output
