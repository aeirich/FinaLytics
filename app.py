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

    tab1, tab2 = st.tabs(['Gesamtübersicht', 'Einzeltitel'])

    with tab1:

        df = Portfolio(uploaded_files)
        df.prepare_df()
        df = df.get_pf()
        st.write(df)

        col1, col2 = st.columns(2)

        with col1:

            PieChart('Einzeltitel',df,'marketvalue','name').plot()

            PieChart('Währungen',df,'marketvalue','ccy').plot()

        with col2:

            BarChart('Performance',df,'performance','name', orientation='vertical').plot()

            PieChart('Asset-Klassen',df,'marketvalue','assetclass').plot()


        HeatmapChart('Heatmap', df, 'assetclass', 'ccy', 'performance').plot()