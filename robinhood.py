from robin_stocks import robinhood
import pyotp

import config

def getStocks():
	totp = pyotp.TOTP(config.get_robinhood_mfa_code()).now()
	login = robinhood.login(config.get_robinhood_username(),config.get_robinhood_password(),mfa_code=totp)

	my_stocks = robinhood.account.build_holdings()

	robinhood.authentication.logout()
	return my_stocks