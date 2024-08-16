def getReturn(ticker, BegDt, EndDt, interval='1d', calc_method='simple'):
	
	"""
	Made by AE
	
	Calculates the return of a stock/index for a given time frame.

	Required arguments:
	
	ticker = str e.g. MSFT
	BegDt = Beginning Timeframe e.g. (2016,1,1)
	EndDt = Ending Timeframe e.g. (2017,12,31)
	interval = Frequency of Data e.g. monthly
		Valid intervals: [1m, 2m, 5m, 15m, 30m, 60m, 90m, 1h, 1d, 5d, 1wk, 1mo, 3mo]
	calc_method = Method of calculation: Logarithm Returns vs. Simple Returns
		Possible values: log=logarithm Returns[Default], simple=simple returns
	Example:

	quote=getReturn('AAPL',(2016,1,1),(2017,12,31), '1m', 'simple')
	print(quote)

	"""
	if calc_method=='simple':
		
		from scipy import stats
		from pandas_datareader import data as pdr
		#import pandas_datareader as pdr
		import numpy as np
		import datetime as dt
		import matplotlib.pyplot as plt
		import datetime as dt
		import yfinance as yf

		yf.pdr_override()
		start_year, start_month, start_day = map(int, BegDt)
		BegDt = dt.datetime(start_year, start_month, start_day)
				
		end_year, end_month, end_day = map(int, EndDt)
		EndDt = dt.datetime(end_year, end_month, end_day)
		
		price=pdr.get_data_yahoo(ticker, BegDt, EndDt, interval=interval)
		Performance = price['Adj Close'].pct_change()
		Performance = Performance.dropna()
		Performance.name='Ret_'+ticker
		
		return Performance
		
	elif calc_method=='log':
		
		from scipy import stats
		from pandas_datareader import data as pdr
		#import pandas_datareader as pdr
		import numpy as np
		import datetime as dt
		import matplotlib.pyplot as plt
		import datetime as dt
		import math as m
		import yfinance as yf

		yf.pdr_override()

		start_year, start_month, start_day = map(int, BegDt)
		BegDt = dt.datetime(start_year, start_month, start_day)
				
		end_year, end_month, end_day = map(int, EndDt)
		EndDt = dt.datetime(end_year, end_month, end_day)
		
		price=pdr.get_data_yahoo(ticker, BegDt, EndDt, interval=interval)
		price.insert(loc=6,column='LogR_'+ticker,value=np.log(price['Adj Close'] / price['Adj Close'].shift(1)))
		price2=price['LogR_'+ticker]
		return price2.dropna()

def getQuote(ticker, BegDt, EndDt, interval='1d'):
	"""
	Made by AE
	
	Returns the quote history of a stock/index for a given time frame.

	Required arguments:
	
	ticker = str e.g. MSFT
	BegDt = Beginning Timeframe e.g. (2016,1,1)
	EndDt = Ending Timeframe e.g. (2017,12,31)
	interval = Frequency of Data e.g. monthly
		Valid intervals: [1m, 2m, 5m, 15m, 30m, 60m, 90m, 1h, 1d, 5d, 1wk, 1mo, 3mo]

	Example:

	quote=getQuote('AAPL',(2016,1,1),(2017,12,31), '1mo')
	print(quote)

	"""
	from scipy import stats
	from pandas_datareader import data as pdr
	#import pandas_datareader as pdr
	import numpy as np
	import datetime as dt
	import matplotlib.pyplot as plt
	import datetime as dt
	import yfinance as yf

	yf.pdr_override()

	start_year, start_month, start_day = map(int, BegDt)
	BegDt = dt.datetime(start_year, start_month, start_day)
			
	end_year, end_month, end_day = map(int, EndDt)
	EndDt = dt.datetime(end_year, end_month, end_day)
	
	price=pdr.get_data_yahoo(ticker, BegDt, EndDt, interval=interval)
	price=price['Adj Close']
	price.name='Q_'+ticker
	return price

def getStatistic(tickerStock, tickerMarket, BegDt, EndDt, statistic, calc_method='simple'):
	"""
	Made by AE
	
	Calculates the entered statistic for a Stock/Market combination for a given time frame.
	Computation is based on monthly returns.

	Required arguments:
	
	tickerStock = Ticker of the Company str e.g. MSFT
	tickerMarket = Ticker of the index str e.g. ^GSPC
	BegDt = Beginning Timeframe e.g. (2016,1,1)
	EndDt = Ending Timeframe e.g. (2017,12,31)
	statistic = beta, corr, p_value, R_squared, std_err, intercept
	calc_method = Calculation is based on either simple or logarythm returns e.g. 'simple', 'log'

	Example:

	beta=getStatistic('AAPL','^GSPC',(2016,1,1),(2017,12,31),'beta', calc_method='log')
	print(beta)

	"""
	
	from scipy import stats
	from pandas_datareader import data as pdr
	#import pandas_datareader as pdr
	import numpy as np
	import datetime as dt
	import matplotlib.pyplot as plt
	import datetime as dt
	import pandas as pd
	import yfinance as yf

	yf.pdr_override()
	
	start_year, start_month, start_day = map(int, BegDt)
	BegDt = dt.datetime(start_year, start_month, start_day)
		
	end_year, end_month, end_day = map(int, EndDt)
	EndDt = dt.datetime(end_year, end_month, end_day)


	if calc_method=='log':
		
		priceStock=pdr.get_data_yahoo(tickerStock, BegDt, EndDt, interval='1d')
		priceStock=priceStock.resample('BM').last()
		priceStock.insert(loc=6,column='Ret_'+tickerStock,value=np.log(priceStock['Adj Close'] / priceStock['Adj Close'].shift(1)))
		LR_Stock=priceStock['Ret_'+tickerStock]
		
		priceMarket=pdr.get_data_yahoo(tickerMarket, BegDt, EndDt, interval='d')
		priceMarket=priceMarket.resample('BM').last()
		priceMarket.insert(loc=6,column='Ret_'+tickerMarket,value=np.log(priceMarket['Adj Close'] / priceMarket['Adj Close'].shift(1)))
		LR_Market=priceMarket['Ret_'+tickerMarket]
	
	elif calc_method=='simple':
		
		priceStock=pdr.get_data_yahoo(tickerStock, BegDt, EndDt, interval='d')
		priceStock=priceStock.resample('BM').last()
		priceStock.insert(loc=6,column='Ret_'+tickerStock,value=priceStock['Adj Close'].pct_change())
		LR_Stock=priceStock['Ret_'+tickerStock]
		
		priceMarket=pdr.get_data_yahoo(tickerMarket, BegDt, EndDt, interval='d')
		priceMarket=priceMarket.resample('BM').last()
		priceMarket.insert(loc=6,column='Ret_'+tickerMarket,value=priceMarket['Adj Close'].pct_change())
		LR_Market=priceMarket['Ret_'+tickerMarket]
	
	
	df=pd.concat([LR_Stock,LR_Market], axis=1)
	df.dropna(inplace=True)
	
	(beta,alpha,r_value,p_value,std_err)=stats.linregress(df['Ret_'+tickerMarket],df['Ret_'+tickerStock])

	if statistic == 'beta':
		return beta
	elif statistic == 'corr':
		return r_value
	elif statistic == 'p_value':
		return p_value
	elif statistic == 'R_squared':
		return r_value**2
	elif statistic == 'std_err':
		return std_err
	elif statistic == 'intercept':
		return alpha

def getVola(ticker, BegDt, EndDt, interval='1mo', calc_method='simple'):
	
	"""
	Made by AE
	
	Calculates the Volatility of a stock/index for a given time frame.

	Required arguments:
	
	ticker = str e.g. MSFT
	BegDt = Beginning Timeframe e.g. (2016,1,1)
	EndDt = Ending Timeframe e.g. (2017,12,31)
	interval = Frequency of Data e.g. monthly
		Valid intervals: [1m, 2m, 5m, 15m, 30m, 60m, 90m, 1h, 1d, 5d, 1wk, 1mo, 3mo]

	Example:

	vola=getVola('AAPL',(2016,1,1),(2017,12,31), '1mo')
	print(vola)

	"""
	if calc_method=='log':
		from scipy import stats
		from pandas_datareader import data as pdr
		#import pandas_datareader as pdr
		import numpy as np
		import datetime as dt
		import matplotlib.pyplot as plt
		import yfinance as yf

		yf.pdr_override()
		start_year, start_month, start_day = map(int, BegDt)
		BegDt = dt.datetime(start_year, start_month, start_day)
				
		end_year, end_month, end_day = map(int, EndDt)
		EndDt = dt.datetime(end_year, end_month, end_day)
		
		price=pdr.get_data_yahoo(ticker, BegDt, EndDt, interval=interval)
		price.insert(loc=6,column='LogR_'+ticker,value=np.log(price['Adj Close'] / price['Adj Close'].shift(1)))
		Performance = price['LogR_'+ticker]
		Performance = Performance.dropna()
		vola = stats.tstd(Performance)
		#vola.name = "Vol_"+ticker
		
		if interval=='d':
			return vola*np.sqrt(250)
		elif interval=='w':
			return vola*np.sqrt(52)
		elif interval =='m':
			return vola*np.sqrt(12)

	elif calc_method=='simple':
		from scipy import stats
		from pandas_datareader import data as pdr
		#import pandas_datareader as pdr
		import numpy as np
		import datetime as dt
		import matplotlib.pyplot as plt
		import yfinance as yf

		yf.pdr_override()

		start_year, start_month, start_day = map(int, BegDt)
		BegDt = dt.datetime(start_year, start_month, start_day)
				
		end_year, end_month, end_day = map(int, EndDt)
		EndDt = dt.datetime(end_year, end_month, end_day)
		
		price=pdr.get_data_yahoo(ticker, BegDt, EndDt, interval=interval)
		price.insert(loc=6,column='Ret_'+ticker,value=price['Adj Close'].pct_change())
		Performance = price['Ret_'+ticker]
		Performance = Performance.dropna()
		vola = stats.tstd(Performance)
		#vola.name = "Vol_"+ticker
		
		if interval=='1d':
			return vola*np.sqrt(250)
		elif interval=='1wk':
			return vola*np.sqrt(52)
		elif interval =='1mo':
			return vola*np.sqrt(12)
			

def getMeanReturn(ticker, BegDt, EndDt, interval='1mo', calc_method='simple'):
	
	"""
	Made by AE
	
	Calculates the geometric Mean Return of a stock/index for a given time frame.

	Required arguments:
	
	ticker = str e.g. MSFT
	BegDt = Beginning Timeframe e.g. (2016,1,1)
	EndDt = Ending Timeframe e.g. (2017,12,31)
	interval = Frequency of Data e.g. monthly
		Possible values: m=monthly[Default], w=weekly, d=daily

	Example:

	mean=getMeanReturn('AAPL',(2016,1,1),(2017,12,31), '1mo')
	print(mean)

	"""
	if calc_method=='log':
		from scipy import stats
		from pandas_datareader import data as pdr
		#import pandas_datareader as pdr
		import numpy as np
		import datetime as dt
		import matplotlib.pyplot as plt
		import datetime as dt
		import yfinance as yf

		yf.pdr_override()

		start_year, start_month, start_day = map(int, BegDt)
		BegDt = dt.datetime(start_year, start_month, start_day)
				
		end_year, end_month, end_day = map(int, EndDt)
		EndDt = dt.datetime(end_year, end_month, end_day)
		
		if interval=='1d':
			price=pdr.get_data_yahoo(ticker, BegDt, EndDt)['Adj Close']
			price=price.ffill()
			MeanReturn = ((np.log(price[-1])-np.log(price[0]))/(len(price)-1))
			return MeanReturn*250
		elif interval=='1wk':
			price=pdr.get_data_yahoo(ticker, BegDt, EndDt)['Adj Close']
			price=price.ffill()
			weekly_prices=price.resample('w').last()
			MeanReturn = ((np.log(weekly_prices[-1])-np.log(weekly_prices[0]))/(len(weekly_prices)-1))
			return MeanReturn*52
		elif interval=='1mo':
			price=pdr.get_data_yahoo(ticker, BegDt, EndDt)['Adj Close']
			price=price.ffill()
			monthly_prices=price.resample('BM').last()
			MeanReturn = ((np.log(monthly_prices[-1])-np.log(monthly_prices[0]))/(len(monthly_prices)-1))
			return MeanReturn*12

	elif calc_method=='simple':
		from scipy import stats
		from pandas_datareader import data as pdr
		#import pandas_datareader as pdr
		import numpy as np
		import datetime as dt
		import matplotlib.pyplot as plt
		import datetime as dt
		import yfinance as yf

		yf.pdr_override()
		
		start_year, start_month, start_day = map(int, BegDt)
		BegDt = dt.datetime(start_year, start_month, start_day)
				
		end_year, end_month, end_day = map(int, EndDt)
		EndDt = dt.datetime(end_year, end_month, end_day)
		

		
		if interval=='1d':
			price=pdr.get_data_yahoo(ticker, BegDt, EndDt)['Adj Close']
			price=price.ffill()
			Performance=price.pct_change()
			MeanReturn=(((1+Performance).product())**(1/(len(Performance)-1)))-1
			return MeanReturn*250
		elif interval=='1wk':
			price=pdr.get_data_yahoo(ticker, BegDt, EndDt)['Adj Close']
			price=price.ffill()
			weekly_prices=price.resample('w').last()
			Performance=weekly_prices.pct_change()
			MeanReturn=(((1+Performance).product())**(1/(len(Performance)-1)))-1
			return MeanReturn*52
		elif interval=='1mo':
			price=pdr.get_data_yahoo(ticker, BegDt, EndDt)['Adj Close']
			price=price.ffill()
			monthly_prices=price.resample('BM').last()
			Performance=monthly_prices.pct_change()
			MeanReturn=(((1+Performance).product())**(1/(len(Performance)-1)))-1
			return MeanReturn*12
		


def getFXhistory(BegDt, EndDt, ForeignCurrency='EUR', DomesticCurrency='EUR'):
	"""
	Made by AE
	
	Downloads a FX Rate history in Base CCY per unit of Investment CCY e.g. EUR/USD
	(Makes performance calculation easier)

	Required arguments:
	
	BegDt = Beginning Timeframe e.g. (2016,1,1)
	EndDt = Ending Timeframe e.g. (2017,12,31)
	ForeignCurrency = Investment CCY e.g. USD
	DomesticCurrency = Reporting CCY = EUR

	Example:

	FX History:

	FX_hist = getFXhistory(EndDt=EndDt, BegDt=BegDt, ForeignCurrency=LocalCCY)

	"""
	try:
		import pandas as pd
		from currency_converter import CurrencyConverter
		from datetime import date, timedelta
		import yfinance as yf

		yf.pdr_override()

		c = CurrencyConverter()
		end_year, end_month, end_day = map(int, EndDt)
		EndDt = date(end_year, end_month, end_day)
		beg_year, beg_month, beg_day = map(int, BegDt)
		BegDt = date(beg_year, beg_month, beg_day)
		delta = EndDt - BegDt
		dates = []
		rates=[]
		for i in range(delta.days + 1):
			day = BegDt + timedelta(days=i)
			x=c.convert(1,DomesticCurrency,ForeignCurrency,day)
			rates.append(x)
			dates.append(day)
		rates_ = pd.Series(rates, index=dates)
		rates_.name = DomesticCurrency+ForeignCurrency
		return rates_
	except Exception as err:
		return 'Rates are only available until the year 2000'
		print('You were trying to download FX Rates as of: ', BegDt)
		print('ERROR: ', err, ': Rates are only available until the year 2000')

def getFXrate(ValDt, ForeignCurrency='EUR', DomesticCurrency='EUR'):
	"""
	Made by AE
	
	Downloads a FX Rate at a certain day in Base CCY per unit of Investment CCY e.g. EUR/USD
	(Makes performance calculation easier)

	Required arguments:
	
	ValDt = Ending Timeframe e.g. (2017,12,31)
	LocalCCY = Investment CCY e.g. USD
	BaseCCY = Reporting CCY = EUR

	Example:

	FX_Rate = getFXrate(EndDt, LocalCCY)
	print(LocalCCY,': ',FX_Rate)


	"""
	try:
		import pandas as pd
		from currency_converter import CurrencyConverter
		from datetime import date, timedelta
		import yfinance as yf

		yf.pdr_override()

		c = CurrencyConverter()
		end_year, end_month, end_day = map(int, ValDt)
		ValDt = date(end_year, end_month, end_day)
		return c.convert(1,DomesticCurrency, ForeignCurrency, ValDt)
	except Exception as err:
		#return 'Rates are only available until the year 2000'
		#return print('You were trying to download FX Rates as of: ', ValDt)
		return print('ERROR: ', err, ': Rates are only available until the year 2000','\n','You were trying to download FX Rates as of: ', ValDt)