import streamlit as st
import pandas as pd

from classes import *

pd.set_option('display.width', 1000)
pd.set_option('display.max_columns', None)
st.set_page_config(page_title='FinaLytics', layout="wide", initial_sidebar_state="auto", menu_items=None)


template = Template().to_excel()


with st.sidebar as sb:

    st.title('FinaLytics `v1.0`')


    st.download_button(
    label="Lade hier dein Template herunter",
    data=template,
    file_name='FinaLytics_Template.xlsx',
    mime='xlsx',
)

    uploaded_files = st.file_uploader("Lade dein Template hoch", type='xlsx')



if uploaded_files:
    df = Portfolio(uploaded_files)
    st.write(df.get_pf())
    df.col_book_value()


#x = Wertpapier('AAPL')
#y=x.get_pricehistory(interval='1d', period='max')
#st.write(y)


