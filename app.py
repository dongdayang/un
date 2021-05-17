import streamlit as st
from load_css import local_css
import pandas as pd
import numpy as np
# Local Imports
import home
import raw_data
import plotting
import reportting
import downloading
import openpyxl

st.set_page_config(layout="wide", page_title='数据可视化 v.0.1')

local_css("style.css")


@st.cache(allow_output_mutation=True)
def load_data(uploaded_file, sheet):
    if uploaded_file is not None:
        excel_file = pd.read_excel(uploaded_file, sheet_name=sheet)

    else:
        excel_file = None

    return excel_file


def load_sheetname(uploaded_file):
    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file, sheet_name=None)

        return list(df)


# Sidebar Options & File Uplaod
excel_file = None
st.sidebar.write('请先上传数据文件')

uploadedfile = st.sidebar.file_uploader(' ', type=['.xlsx'], help='若改变功能，请重新上传文件')

# Sidebar Navigation
st.sidebar.title('导航')

options = st.sidebar.radio('选择:',
                           ['主页', '表预览', '图表分析', '自动生成周报'])

if options == '主页':
    home.home()
elif options == '表预览':
    if load_sheetname(uploadedfile) is not None:
        state_selected = st.selectbox(
            '选择sheet',
            load_sheetname(uploadedfile),
        )
        excel_file = load_data(uploadedfile, state_selected)
        raw_data.raw_data(excel_file)

elif options == '图表分析':
    if load_sheetname(uploadedfile) is not None:
        sheet_selected = st.selectbox(
            '选择sheet',
            load_sheetname(uploadedfile),
        )
        excel_file = load_data(uploadedfile, sheet_selected)
        plotting.plot(excel_file,sheet_selected)
elif options == '自动生成周报':
    state_selected = st.selectbox(
        '预览',
        ['未选择', '组织区域满编率', '校区满编率', '组织区域入职离职数据', '校区入职离职数据'],
    )
    excel_file = load_data(uploadedfile, '在职')
    excel_file2 = load_data(uploadedfile, '入职')
    excel_file3 = load_data(uploadedfile, '离职')
    if load_sheetname(uploadedfile) is not None:
        reportting.report(excel_file, excel_file2, excel_file3, state_selected)
