import yfinance as yf
from io import BytesIO
import pandas as pd

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
        self.name = None
        self.data = None
        self.assettyp = None
        self.load_yf_data()
        self.get_base_data()

    def load_yf_data(self):
        import yfinance as yf
        self.data = yf.Ticker(self.ticker)

    def get_base_data(self):
        self.isin = self.data.isin
        self.name = self.data.info['shortName']
        self.sector = self.data.info['sector']
        self.industry = self.data.info['industry']
        self.country = self.data.info['country']
        self.ccy = self.data.info['currency']
        self.longBusinessSummary = self.data.info['longBusinessSummary']
    
    def get_pricehistory(self, interval='1mo', period='10y'):
        """
        params: 
            interval (str): 1m, 2m, 5m, 15m, 30m, 60m, 90m, 1h, 1d, 5d, 1wk, 1mo, 3mo
            period (str): 1d, 5d, 1mo, 3mo, 6mo, 1y, 2y, 5y, 10y, ytd, max
            startDate (datetime.)
        """
        self.pricehistory = self.data.history(interval=interval,period=period)['Close']
        return self.pricehistory


class Template:
    def __init__(self):
        import pandas as pd
        self.data = pd.DataFrame({
            'Yahoo Ticker':['AAPL'],
            'Kaufpreis':[140.43],
            'Bestand':[30],
            'Kaufdatum':['21.03.2023']
            },index=None)



    def to_excel(self):
        # Speichere das DataFrame in einen BytesIO-Stream
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            self.data.to_excel(writer, index=False)
        return output.getvalue()


class Portfolio:
    def __init__ (self, data):
        self.data = pd.read_excel(data)
        self.pf = self.data
        print('Excel file eingelesen')
        print(self.data.head())
        self.geomean_return = None
        self.portfolio_ccy = None

    def col_book_value(self):
        if 'Kaufpreis' in self.pf.columns and 'Bestand' in self.pf.columns:
            self.pf['book_value'] = self.pf['Kaufpreis'] * self.pf['Bestand']
            print("Spalte 'Marktwert' hinzugefügt.")
        else:
            print("Die erforderlichen Spalten 'Price' und 'Menge' fehlen.")


    def geo_mean_return(self, pct_change_column):
        
        self.geomeanreturn = (((1 + pct_change_column).product()) ** (1 / (len(pct_change_column))) - 1) * 12
        return self.geomeanreturn

    def get_pf(self):
        """
        Zeigt das Porfolio bereinigt um nicht benötigte und hinzugefügte Spalten
        """
        return self.pf


class Charts:
    def __init__(self, title, data):
        self.title = title
        self.data = data
    
    def display(self):
        print(f"Anzeigen des Charts: {self.title}")
    
    def update_data(self, new_data):
        self.data = new_data

class PieChart(Charts):
    def __init__(self, title, data, colors):
        super().__init__(title, data)
        self.colors = colors
    
    def display(self):
        super().display()
        print("Anzeigen als Kreisdiagramm")

class BarChart(Charts):
    def __init__(self, title, data, orientation='vertical'):
        super().__init__(title, data)
        self.orientation = orientation
    
    def display(self):
        super().display()
        print(f"Anzeigen als Balkendiagramm mit {self.orientation} Ausrichtung")
