import pandas as pd

def read_excel_file(uploaded_file):
    try:
        excel_data = pd.read_excel(uploaded_file, sheet_name=None)  # 读取所有工作表
        return excel_data
    except Exception as e:
        return {"错误": pd.DataFrame({"错误信息": [str(e)]})}
