import streamlit as st


def raw_data(excel_file):
    st.dataframe(excel_file)
