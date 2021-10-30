import datetime
import logging
import sys
import traceback
import pandas

import excel
import robinhood
import config
import archive

def main():
	stocks = robinhood.getStocks()
	excel.loadIntoExcel(stocks)
	writeToCSV(stocks)
	archive.archiveAll()

def writeToCSV(stocks):
	df = pandas.read_csv(config.get_csv_fileName())
	todayString = datetime.datetime.now().strftime("%m/%d/%Y %H:%M:%S")
	for stock in stocks:
		stockEquity = float(stocks[stock]['equity'])
		stockAmountInvested = stockEquity - float(stocks[stock]["equity_change"])
		listToAdd = [todayString,stock,round(stockAmountInvested,2),round(stockEquity,2)]
		add_series = pandas.Series(listToAdd,index = df.columns)

		df = df.append(add_series,ignore_index = True)

	df.to_csv(config.get_csv_fileName(),index=False)


if __name__ == "__main__":
	try:
		main()
	except Exception as e:
		exc_type, exc_obj, exc_tb = sys.exc_info()
		logging.info("\t\t" + str(e) + " at " + datetime.datetime.now().strftime("%m/%d/%Y %H:%M:%S:%f"))
		logging.info("Call Stack:")
		callStack = traceback.format_exc()
		logging.info( callStack + "\n\n")