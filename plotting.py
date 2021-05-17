import streamlit as st
import pandas as pd


def plot(excel_file,sheet_selected):
    st.title('招聘数据分析')
    state_selected = st.selectbox(
        '选择数据',
        ['组织区域', '部门\校区'],
    )
    if excel_file is not None:
        df = excel_file[state_selected].value_counts()
        st.write(sheet_selected+state_selected + "招聘直方图")
        st.bar_chart(df)
