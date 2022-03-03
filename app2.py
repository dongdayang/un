import streamlit as st
import plotting2
import pandas as pd
import home
import plotly
import raw_data

st.set_page_config(layout="wide", page_title='数据分析 v.0.1')


@st.cache(allow_output_mutation=True)
def load_data(uploaded_file, sheet):
    if uploaded_file is not None:
        if typename == 'xlsx':
            excelfile = pd.read_excel(uploaded_file, sheet_name=sheet)
        elif typename == 'csv':
            excelfile = pd.read_csv(uploaded_file, encoding='GB18030')

    else:
        excelfile = None

    return excelfile


@st.cache
def load_sheetname(uploaded_file):
    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file, sheet_name=None)

        return list(df)


@st.cache
def load_indexname(excelfile):
    if excelfile is not None:
        df = excelfile
        name = df.columns.values

        return name


# Sidebar Options & File Upload
excel_file = None
index_selected = None
typename = None
sheetname = None
st.sidebar.write('请先上传数据文件')

uploadedfile = st.sidebar.file_uploader(' ', type=['.xlsx', 'csv'], help='若改变功能，请重新上传文件')
if uploadedfile is not None:
    typename = uploadedfile.name.split('.')[-1]
    if typename == 'xlsx':
        sheetname = load_sheetname(uploadedfile)

# Sidebar Navigation
st.sidebar.title('导航')

options = st.sidebar.radio('选择:',
                           ['主页', '表预览', '图表分析'])

if options == '主页':
    home.home()
elif options == '表预览':
    if typename == 'xlsx':
        if sheetname is not None:
            state_selected = st.selectbox(
                '选择sheet',
                sheetname,
            )
            excel_file = load_data(uploadedfile, state_selected)
            raw_data.raw_data(excel_file)

    elif typename == 'csv':
        excel_file = load_data(uploadedfile, None)
        raw_data.raw_data(excel_file)


elif options == '图表分析':
    if typename == 'xlsx':
        if sheetname is not None:
            sheet_selected = st.selectbox(
                '选择sheet',
                sheetname,
            )

            excel_file = load_data(uploadedfile, sheet_selected)
            indexname=load_indexname(excel_file)
            x_selected = st.selectbox(
                '选择x轴',
                indexname,
            )
            index_selected = st.multiselect(
                '选择数据',
                indexname,
            )
            plotting2.plot(excel_file, x_selected, index_selected)
    elif typename == 'csv':
        excel_file = load_data(uploadedfile, None)
        indexname = load_indexname(excel_file)
        x_selected = st.selectbox(
            '选择x轴',
            indexname,
        )
        index_selected = st.multiselect(
            '选择数据',
            indexname,
        )
        plotting2.plot(excel_file, x_selected, index_selected)
        # , index_selected , excel_file2, excel_file3, sheet_selected)
