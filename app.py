import streamlit as st
import pandas as pd

pd.set_option('display.width', 1000)
pd.set_option('display.max_columns', None)
st.set_page_config(page_title='FinaLytics', layout="wide", initial_sidebar_state="auto", menu_items=None)

with st.sidebar as sb:

    st.title('FinaLytics `v1.0`')


    uploaded_files = st.file_uploader("Lade dein Template hoch", type='xlsx')

if uploaded_files:
    df = pd.read_excel(uploaded_files)
    st.write(df)