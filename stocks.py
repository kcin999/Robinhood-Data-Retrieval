import robin_stocks 
import pyotp
import yfinance

import config

def getRobinhoodStocks():
	totp = pyotp.TOTP(config.get_robinhood_mfa_code()).now()
	login = robin_stocks.robinhood.login(config.get_robinhood_username(),config.get_robinhood_password(),mfa_code=totp)

	my_stocks =robin_stocks. robinhood.account.build_holdings()

	temp = robin_stocks.robinhood.get_open_stock_positions()

	robin_stocks.robinhood.authentication.logout()
	return my_stocks

def getPastData(ticker):
	data = yfinance.download(tickers=ticker,period = "1d", interval = "1h")
	return data