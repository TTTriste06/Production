import streamlit as st

def setup_sidebar():
    with st.sidebar:
        st.title("📊 Excel 文件预览工具")
        st.markdown("上传一个 Excel 文件并查看其内容。")

def upload_excel_file():
    st.header("🗓️ Excel 处理")
    
    uploaded_file = st.file_uploader("上传 Excel 文件", type=["xlsx"])
    return uploaded_file
