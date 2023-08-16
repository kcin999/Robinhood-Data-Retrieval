#Module Imports
import os
import robin_stocks 
import pyotp
import yfinance
import pandas as pd
import time

#File Imports
import config

def getRobinhoodStocks():
	totp = pyotp.TOTP(os.environ.get('ROBINHOOD_ONETIME_PASSWORD')).now()
	_login = robin_stocks.robinhood.login(os.environ.get('ROBINHOOD_USERNAME'),os.environ.get('ROBINHOOD_PASSWORD'),mfa_code=totp)

	my_stocks =robin_stocks.robinhood.account.build_holdings()

	robin_stocks.robinhood.authentication.logout()
	return my_stocks

def getPastData(ticker):
	data = yfinance.download(tickers=ticker,period = "1d", interval = "1h")
	return data

def getRobinhoodTransactions():
	totp = pyotp.TOTP(os.environ.get('ROBINHOOD_ONETIME_PASSWORD')).now()
	_login = robin_stocks.robinhood.login(os.environ.get('ROBINHOOD_USERNAME'),os.environ.get('ROBINHOOD_PASSWORD'),mfa_code=totp)

	# Option Orders
	option_orders =robin_stocks.robinhood.get_all_stock_orders()
	option_orders = pd.DataFrame.from_records(option_orders)
	option_orders[['Name', 'Simple Name', 'Symbol']] = option_orders.apply(lambda x: getRobinhoodInstrument(x['instrument']), result_type='expand', axis=1)
	option_orders = option_orders[['Name', 'Simple Name', 'Symbol', 'price', 'quantity', 'state', 'side', 'last_transaction_at']]


	# Dividends
	dividends = robin_stocks.robinhood.get_dividends()
	dividends = pd.DataFrame.from_records(dividends)
	dividends[['Name', 'Simple Name', 'Symbol']] = dividends.apply(lambda x: getRobinhoodInstrument(x['instrument']), result_type='expand', axis=1)

	dividends = dividends[['Name', 'Simple Name', 'Symbol', 'state', 'amount', 'payable_date']]

	robin_stocks.robinhood.authentication.logout()


	return option_orders, dividends


def getRobinhoodInstrument(url: str):
	instrument = robin_stocks.robinhood.get_instrument_by_url(url)

	return instrument['name'], instrument['simple_name'], instrument['symbol']
