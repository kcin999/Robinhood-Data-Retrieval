#Module Imports
import pandas
import datetime
import stocks

#File Imports
import config


def writeToCSV(stocks):
	updateMyStockData(stocks)
	updateMarketData()

def updateMyStockData(stocks):
	df = pandas.read_csv(config.get_csv_myStocksFileName())
	todayString = datetime.datetime.now().strftime("%m/%d/%Y %H:%M:00")
	for stock in stocks:
		stockEquity = float(stocks[stock]['equity'])
		stockAmountInvested = stockEquity - float(stocks[stock]["equity_change"])
		listToAdd = [todayString,stock,round(stockAmountInvested,2),round(stockEquity,2)]
		add_series = pandas.Series(listToAdd,index = df.columns)

		df = df.append(add_series,ignore_index = True)

	df = df.sort_values(['Date','Stock'], ascending=[0,1])
	df.to_csv(config.get_csv_myStocksFileName(),index=False)

def updateMarketData():
	#Finds the tickers of all stocks I own from my personal stock data csv. List(set(...)) are to remove duplicates from list
	stocksIHaveOwned = list(set(pandas.read_csv(config.get_csv_myStocksFileName())['Stock'].to_list()))
	df = pandas.read_csv(config.get_csv_marketStockDataFileName())
	for stock in stocksIHaveOwned:
		stockDF = stocks.getPastData(stock)

		#Iterate through DF
		for i in range(0,len(stockDF)):
			#Finds the date that the data is for
			dateToAdd = stockDF.index[i].strftime('%m/%d/%Y %H:%M:00')
			#Stock Price at close for the period
			stockPrice = round(stockDF.values[i][3],2)

			#Appends it to CSV
			listToAdd = [dateToAdd,stock,stockPrice]
			add_series = pandas.Series(listToAdd,index = df.columns)
			df = df.append(add_series,ignore_index = True)

	df = df.drop_duplicates(subset=['date','stock'],keep='last')
	df = df.sort_values(['date','stock'],ascending=[0,1])
	df.to_csv(config.get_csv_marketStockDataFileName(),index=False)