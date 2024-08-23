import yfinance as yf

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
        self.isin = None
        self.assetclass = None
        self.name = None
        self.currency = None


class Template:
    def __init__ (self):
        self.excel = None

class Portfolio:
    def __init__ (self, data):
        self.data = data
        self.geomean_return = None
        self.portfolio_ccy = None

    def geo_mean_return(self, pct_change_column):
        
        self.geomeanreturn = (((1 + pct_change_column).product()) ** (1 / (len(pct_change_column))) - 1) * 12
        return self.geomeanreturn

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
