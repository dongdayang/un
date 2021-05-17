import streamlit as st
import pandas as pd

state_selected = None


def plot(excel_file, excel_file2, excel_file3, sheet_selected):
    st.title('招聘数据分析')
    state_selected = st.selectbox(
        '选择数据',
        ['组织区域', '部门\校区'],
    )
    if state_selected is not None:
        df = excel_file[state_selected].value_counts()
        st.write(sheet_selected + state_selected + "招聘直方图")
        st.bar_chart(df)
        df1 = excel_file2[state_selected].value_counts()
        df2 = -excel_file3[state_selected].value_counts()
        df3 = pd.concat([df1, df2])
        st.write("入职离职招聘直方图")
        st.bar_chart(df3)
