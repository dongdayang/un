import streamlit as st
import plotly.graph_objs as go

@st.cache(suppress_st_warning=True)
def plot(excel_file, x_selected, index_selected):
    list = []
    for i in index_selected:
        chart = go.Scatter(
            x=excel_file[x_selected],
            y=excel_file[i],
            mode='lines+markers',
            name=i
        )

        list.append(chart)
    if list:
        st.plotly_chart(list,use_container_width=True)
