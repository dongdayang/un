import streamlit as st
import pandas as pd

# Plotly imports
import matplotlib.pyplot as plt
import numpy as np


def plot(excel_file):
    st.title('季度数据')
    month_selected = st.slider("选择月份", 1, 12)
    if excel_file is not None:
        st.write(month_selected)
        df = excel_file['组织区域'].value_counts()
        st.write("各大区招聘直方图")
        st.bar_chart(df)
