import yfinance as yf
from io import BytesIO
import pandas as pd
import plotly.express as px
import streamlit as st
import numpy as np
import plotly.graph_objects as go
from currency_converter import CurrencyConverter
from datetime import timedelta
import sys

def colorset():
    color_dict = {
       'darkorange':'rgb(255,140,0)',
       'darkorange4':'rgb(139,69,0)',
       'gold':'rgb(255,215,0)',
       'PeachPuff':'rgb(255,218,185)',
       'honeydew3':'rgb(193,205,193)',
       'lightpink3':'rgb(205,140,149)',
       'orangered':'rgb(255,69,0)',
       'gold':'rgb(255,215,0)'}
    colors = list(color_dict.values())
    return colors


class Template:
    def __init__(self):
        import pandas as pd
        self.data = pd.DataFrame({
            'Yahoo Ticker':['AAPL','XS2689948078.SG','MSE.PA'],
            'Name (Optional)':['Apple Inc.','6,375% Rumänien 23/33','Amundi EURO STOXX 50 II UCITS ETF Acc'],
            'Kaufpreis':[140.43, 150.32,100],
            'Bestand':[30,1000,100],
            'Kaufdatum':['21.03.2023','21.03.2023','21.03.2022'],
            'Anlageklasse':['Aktien','Anleihen','Aktien-ETF']},index=None)

    def to_excel(self):
        # Speichere das DataFrame in einen BytesIO-Stream
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            self.data.to_excel(writer, index=False)
        return output.getvalue()
    


class Wertpapier:
    """Sammelt die Attribute und Methoden von Wertpapieren
    
    Die Klasse soll die Attribute z.B. Ticker, ISIN, Asset-Klasse in sich vereinen.
    
    Args:
        ticker (str): Yahoo Ticker zum Download von Daten.
        isin (str): Offizelle ISIN (International Securities Identification Number)
        assetclass (str): Gibt die Klassifizierung als Aktie, Rente, Aktienfonds, Rentenfonds, Mischfonds, Aktien-ETF, Renten-ETF, Kryptowährungen
        name (str): Name des Wertpapiers
        currency (str): Gibt die Währung des Instruments wieder.

    """
    def __init__ (self, ticker):
        self.ticker = ticker
        self.assetclass = None
        self.data = None
        self.assettyp = None
        self.load_yf_data()
        self.get_base_data()

    def load_yf_data(self):
        import yfinance as yf
        self.data = yf.Ticker(self.ticker)

    def get_base_data(self):
        self.isin = self.data.isin if self.data.isin else None
        self.name = self.data.info.get('shortName',None)
        self.longname = self.data.info.get('longName',None)
        self.sector = self.data.info.get('sector',None)
        self.industry = self.data.info.get('industry',None)
        self.country = self.data.info.get('country')
        self.ccy = self.data.info.get('currency')
        self.longBusinessSummary = self.data.info.get('longBusinessSummary',None)
        self.price = self.data.info.get('previousClose',None)
        self.beta = self.data.info.get('beta',None)
    
    def get_pricehistory(self, interval='1mo', period='10y'):
        """
        params: 
            interval (str): 1m, 2m, 5m, 15m, 30m, 60m, 90m, 1h, 1d, 5d, 1wk, 1mo, 3mo
            period (str): 1d, 5d, 1mo, 3mo, 6mo, 1y, 2y, 5y, 10y, ytd, max
            startDate (datetime.)
        """
        self.pricehistory = pd.DataFrame(self.data.history(interval=interval,period=period)['Close'])
        #self.pricehistory.index = self.pricehistory.index.strftime('%Y-%m-%d')
        return self.pricehistory


class Portfolio:
    def __init__ (self, data):
        self.data = pd.read_excel(data)
        self.pf = self.data
        print('Excel file eingelesen')
        print(self.data.head())
        self.geomean_return = None
        self.portfolio_ccy = 'EUR'


    def prepare_df(self):
            
        self.pf.rename(columns={'Yahoo Ticker':'ticker','Name (Optional)':'name','Kaufdatum':'purchasedate','Kaufpreis':'bookprice', 'Bestand':'amount', 'Anlageklasse':'assetclass'},inplace=True)
        self.pf['purchasedate'] = pd.to_datetime(self.pf['purchasedate']).dt.date
        self.wertpapiere = [Wertpapier(ticker) for ticker in self.pf['ticker']]

        self.pf['isin'] = [wp.isin for wp in self.wertpapiere]
        self.pf['name'] = [self.pf['name'][i] if i < len(self.pf['name']) else wp.name for i, wp in enumerate(self.wertpapiere)]
        self.pf['current_price'] = [wp.price for wp in self.wertpapiere]
        self.pf['ccy'] = [wp.ccy for wp in self.wertpapiere]
        self.pf['industry'] = [wp.industry for wp in self.wertpapiere]
        self.pf['beta'] = [wp.beta for wp in self.wertpapiere]
        self.pf['country'] = [wp.country for wp in self.wertpapiere]
        
        # Berechnung des aktuellen Werts
        self.pf['bookvalue'] = np.where(self.pf['ticker'].str.contains('.SG'),self.pf['bookprice'] * self.pf['amount'] * 0.01, self.pf['bookprice'] * self.pf['amount'])
        self.pf['marketvalue'] = np.where(self.pf['ticker'].str.contains('.SG'),self.pf['current_price'] * self.pf['amount'] * 0.01, self.pf['current_price'] * self.pf['amount'])

        # Berechnung der Performance
        self.pf['performance'] = (self.pf['marketvalue'] - self.pf['bookvalue']) / self.pf['bookvalue']
        self.pf['weights'] = (self.pf['bookvalue'] / self.pf['bookvalue'].sum())

    def key_statistics(self):
        try:
            self.marketvalue = self.pf['marketvalue'].sum()
            self.bookvalue = self.pf['bookvalue'].sum()
            self.totalreturn = self.pf['performance'].dot(self.pf['weights']).sum()
        except:
            return 'Führe zuerst Porfolio.prepare_df aus'

    def geo_mean_return(self, pct_change_column):
        
        self.geomeanreturn = (((1 + pct_change_column).product()) ** (1 / (len(pct_change_column))) - 1) * 12
        return self.geomeanreturn

    def get_pf(self):
        """
        Zeigt das Porfolio bereinigt um nicht benötigte und hinzugefügte Spalten
        """
        
        return self.pf

    def get_quotes(self, period='10y', foreignCurrency=False):
        """
        params:
        foreignCurrency (boolean): False = Download in EUR / True = Download in Original Currency (FC = Foreign Currency / DC = Domestic Currency)
                Defines whether the quotes should be downloaded in Original Currency oder in EUR.
        period (str): 1d, 5d, 1mo, 3mo, 6mo, 1y, 2y, 5y, 10y, ytd, max
        """
        from datetime import datetime 

        quotes = pd.DataFrame()
        tmp_min_dates = pd.DataFrame()
        min_dates=[]
        print('Daten werden heruntergeladen und aufbereitet....')

        for ticker in self.pf['ticker']:
            try:
                i = Wertpapier(ticker)
                price_history = i.get_pricehistory(interval='1d', period=period)
                price_history.index = pd.to_datetime(price_history.index).date
                price_history = price_history.rename(columns={'Close': i.ticker})

                if quotes.empty:
                    quotes = price_history
                else:
                    quotes = quotes.join(price_history, how='outer')
            except Exception as err:
                print('Error: ', err)
                print('Security: ', i.name)
                continue
        quotes.index = pd.to_datetime(quotes.index)
        monthly_quotes_FC = quotes.resample('BME').last()

        if foreignCurrency==False:

            #-------------- Download Währungen --------------------------------

            currencies = self.pf['ccy'].drop_duplicates()
            c = CurrencyConverter()
            df_CCY = pd.DataFrame()

            for ccy in currencies:
                fx_rates = []
                for day in monthly_quotes_FC.index.date:
                    success = False
                    try_day = day  # Starte mit dem aktuellen Tag
                    while not success:
                        try:
                            x = c.convert(1, 'EUR', ccy, try_day)
                            fx_rates.append(x)
                            success = True  # Wert erfolgreich gefunden, Schleife verlassen
                        except:
                            try_day -= timedelta(days=1)  # Einen Tag zurückgehen und erneut versuchen
                            if try_day < monthly_quotes_FC.index.date[0]:  # Vermeide unendliche Schleifen
                                fx_rates.append(None)  # Kein gültiger Wert gefunden
                                success = True
                                break

                fx_rates_ = pd.Series(fx_rates, index=monthly_quotes_FC.index.date)
                df_CCY.insert(loc=len(df_CCY.columns),column=ccy,value=fx_rates)

            df_CCY.index=monthly_quotes_FC.index.date

            # --------------Calculate Domestic Return ----------------------
            monthly_quotes_DC = pd.DataFrame()
            monthly_quotes_data = monthly_quotes_FC.join(df_CCY,how='outer')
            for i in self.pf['ticker']:
                ccy = (self.pf.loc[self.pf['ticker']==i,['ccy']].values)[0][0]
                try:
                    monthly_quotes_DC[str(i)] = monthly_quotes_data[i]/monthly_quotes_data[ccy]
                except Exception as err:
                    print('Error: ', err)
                    continue

        print('Daten verfügbar....')

        return monthly_quotes_DC if foreignCurrency==False else monthly_quotes_FC



class Charts:
    def __init__(self, title, data):
        self.title = title
        self.data = data
        self.colorset = colorset()

    def colors(self, values):
        """Gibt eine Liste von Farben basierend auf den Werten zurück."""
        return ['green' if x >= 0 else 'red' for x in values]

    def plot(self):
        raise NotImplementedError("Subklassen müssen die plot-Methode implementieren.")
    
    def update_data(self, new_data):
        self.data = new_data

class PieChart(Charts):
    def __init__(self, title, data, category, values):
        super().__init__(title, data)
        self.values = values #TODO testen in dekalytics df[col] = df[col].clip(lower=0)
        self.category = category
    
    def plot(self):
        # Erstellen des Kreisdiagramms
        fig = px.pie(self.data, values=self.values, names=self.category, title=self.title, color_discrete_sequence=self.colorset)
        fig.update_layout(width=800, height=600)
        st.plotly_chart(fig)

        print("Anzeigen als Kreisdiagramm")

class BarChart(Charts):
    def __init__(self, title, data, category, values, orientation='vertical'):
        super().__init__(title, data)
        self.values = values
        self.category = category
        self.orientation = orientation

    def plot(self):
        # Farben auf die Daten anwenden
        colors = self.colors(self.data[self.values])
        
        # Erstellen des Balkendiagramms
        fig = px.bar(self.data,
                     y=self.values if self.orientation == 'vertical' else self.category,
                     x=self.category if self.orientation == 'vertical' else self.values,
                     title=self.title)
        fig.update_traces(marker=dict(color=colors))
        width = 800 if self.orientation == 'vertical' else 600
        height = 800 if self.orientation == 'horizontal' else 600
        fig.update_layout(width=width, height=height, xaxis={'categoryorder': 'total descending'})
        st.plotly_chart(fig)
        
        print(f"Anzeigen als Balkendiagramm mit {self.orientation} Ausrichtung")


class HeatmapChart(Charts):
    def __init__(self, title, data, x_category, y_category, values):
        super().__init__(title, data)
        self.x_category = x_category
        self.y_category = y_category
        self.values = values
        self.grouped_data = None

    def prepare_data(self, aggregation='mean', bins=None):
        if bins:
            self.data[f'{self.x_category}_bin'] = pd.cut(self.data[self.x_category], bins=bins)
            self.data[f'{self.y_category}_bin'] = pd.cut(self.data[self.y_category], bins=bins)
            x_category, y_category = f'{self.x_category}_bin', f'{self.y_category}_bin'
        else:
            x_category, y_category = self.x_category, self.y_category

        grouped = self.data.groupby([x_category, y_category])[self.values].agg(aggregation).reset_index()
        self.grouped_data = grouped.pivot(index=y_category, columns=x_category, values=self.values)

    def plot(self):
        if self.grouped_data is None:
            self.prepare_data()

        fig = go.Figure(data=go.Heatmap(
            z=self.grouped_data.values,
            x=self.grouped_data.columns,
            y=self.grouped_data.index,
            colorscale='RdYlGn',
            hoverongaps=True
        ))

        fig.update_layout(
            title=self.title,
            xaxis_title=self.x_category,
            yaxis_title=self.y_category,
            width=800,
            height=600
        )

        st.plotly_chart(fig)

        print("Anzeigen als Heatmap")


class IcicleChart(Charts):
    def __init__(self, title, data, list_categories, values):
        super().__init__(title, data)
        self.values = values #TODO testen in dekalytics df[col] = df[col].clip(lower=0)
        self.category = list_categories
    
    def plot(self):
        print(self.category)
        constant = [px.Constant("Gesamt")]
        #self.category = self.category.insert(0,px.Constant("Gesamt"))
        self.category = constant + self.category
        print(self.category)
        print(self.values)
        fig = px.icicle(self.data, values=self.values, path=self.category, title=self.title, color_discrete_sequence=self.colorset)
        fig.update_traces(root_color="lightgrey")
        fig.update_layout(width=800, height=600)
        st.plotly_chart(fig)

        print("Anzeigen als Icicle")