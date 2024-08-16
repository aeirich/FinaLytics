import streamlit as st
import pandas as pd

pd.set_option('display.width', 1000)
pd.set_option('display.max_columns', None)
st.set_page_config(page_title='ImmoLytics',page_icon=":house_buildings:", layout="wide", initial_sidebar_state="auto", menu_items=None)
st.title('FinaLytics `v1.0`')


uploaded_files = st.file_uploader(
    "Choose a CSV file", accept_multiple_files=True
)
for uploaded_file in uploaded_files:
    bytes_data = uploaded_file.read()
    st.write("filename:", uploaded_file.name)
    st.write(bytes_data)