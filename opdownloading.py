import streamlit as st
import base64
from io import BytesIO
import pandas as pd
import openpyxl
from openpyxl.writer.excel import save_virtual_workbook


def to_excel(df):

    processed_data = save_virtual_workbook(df)
    return processed_data


def get_table_download_link(df, state_selected,name):
    val = to_excel(df)
    b64 = base64.b64encode(val)
    return f'<a href="data:application/octet-stream;base64,{b64.decode()}" download="'+state_selected+'.xlsx">下载'+name+'</a>'
