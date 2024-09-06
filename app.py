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

    pf = Portfolio(uploaded_files)
    pf.prepare_df()
    pf.key_statistics()

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.metric(label='Gesamtwert: ', value=f'{round(pf.marketvalue, 2)} €', delta=(f'{round((pf.marketvalue - pf.bookvalue),2)} €'), delta_color='normal')

    with col2:
        st.metric(label = 'Total Return', value=f'{round(pf.totalreturn*100,2)} %')

    tab1, tab2 = st.tabs(['Gesamtübersicht', 'Einzeltitel'])

    with tab1:

        df = pf.get_pf()

        st.write(df)

        quotes = pf.get_quotes(period='1y', foreignCurrency=False)

        st.write(quotes)

        col1, col2 = st.columns(2)

        with col1:

            PieChart('Einzeltitel',df,values='marketvalue',category='name').plot()

            PieChart('Währungen',df,'ccy','marketvalue').plot()

            HeatmapChart('Heatmap', df, 'assetclass', 'ccy', 'performance').plot()

        with col2:

            BarChart('Performance',df,'name','performance', orientation='vertical').plot()

            PieChart('Asset-Klassen',df,'assetclass','marketvalue').plot()

            path = ['ccy', 'assetclass', 'name']
        
            IcicleChart('Icicle Test', df, path, 'marketvalue').plot()


        LineChart('Line Chart',quotes,y_values=quotes.columns,x_values=quotes.index).plot()
