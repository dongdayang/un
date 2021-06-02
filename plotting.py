import matplotlib.pyplot as plt
import streamlit as st
import matplotlib as mpl
import matplotlib.font_manager as fm

fe = fm.FontEntry(
    fname='./Microsoft-YaHei.tff',
    name='Microsoft-YaHei')
fm.fontManager.ttflist.insert(0, fe) # or append is fine
plt.rcParams['font.family'] = fe.name # = 'your custom ttf font name'
plt.rcParams['axes.unicode_minus'] = False  # 用来正常显示负号

state_selected = None

st.set_option('deprecation.showPyplotGlobalUse', False)


def plot(excel_file, excel_file2, excel_file3, sheet_selected):
    state = st.selectbox(
        '选择数据',
        ['组织区域', '部门\校区'],
    )
    if state is not None:
        df1 = excel_file2[state].value_counts()
        df2 = excel_file3[state].value_counts()
        x = df1.index
        x1 = df2.index
        y = df1
        y1 = df2

        plt.bar(x, y, align='center')
        plt.bar(x1, y1, color='orange', align='center')
        plt.legend(labels=['入职', '离职'])  # 图例
        plt.title('入职离职招聘直方图')
        plt.ylabel('入离职数据')
        plt.xlabel(state)
        plt.xticks(rotation=90)

        st.pyplot()

        df = excel_file[state].value_counts()
        st.write(sheet_selected + state + "招聘直方图")
        st.bar_chart(df)
