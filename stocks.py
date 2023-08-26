"""Stock Functions"""
# Module Imports
import os
import robin_stocks
import pyotp
import yfinance
import pandas as pd


def get_robinhood_stocks() -> dict:
    """Gets Robinhood Stock information from my account

    :return: Robinhood Stock information
    :rtype: dict
    """
    totp = pyotp.TOTP(os.environ.get('ROBINHOOD_ONETIME_PASSWORD')).now()
    _login = robin_stocks.robinhood.login(
        os.environ.get('ROBINHOOD_USERNAME'),
        os.environ.get('ROBINHOOD_PASSWORD'),
        mfa_code=totp
    )

    my_stocks = robin_stocks.robinhood.account.build_holdings()

    robin_stocks.robinhood.authentication.logout()
    return my_stocks


def get_historical_data(ticker: str, period: str = '60d', interval: str = '5m') -> pd.DataFrame:
    """Gets historical data for a given stock ticker.

    :param ticker: Ticker to get data for
    :type ticker: str
    :param period: Period to get data for, defaults to '60d'
    :type period: str, optional
    :param interval: Time between records, defaults to '5m'
    :type interval: str, optional
    :return: _description_
    :rtype: pd.DataFrame
    """

    data = yfinance.download(tickers=ticker, period=period, interval=interval)
    return data


def get_robinhood_transactions() -> tuple[pd.DataFrame, pd.DataFrame]:
    """Returns Robinhood Transactions from the user account

    :return: Transactions and Divdend dataframes
    :rtype: tuple[pd.DataFrame, pd.DataFrame]
    """
    totp = pyotp.TOTP(os.environ.get('ROBINHOOD_ONETIME_PASSWORD')).now()
    _login = robin_stocks.robinhood.login(
        os.environ.get('ROBINHOOD_USERNAME'),
        os.environ.get('ROBINHOOD_PASSWORD'),
        mfa_code=totp
    )

    # Option Orders
    option_orders = robin_stocks.robinhood.get_all_stock_orders()
    option_orders = pd.DataFrame.from_records(option_orders)
    option_orders[['Name', 'Simple Name', 'Symbol']] = option_orders.apply(
        lambda x: get_robinhood_instrument(x['instrument']), result_type='expand', axis=1)
    option_orders = option_orders[['Name', 'Simple Name', 'Symbol',
                                   'price', 'quantity', 'state', 'side', 'last_transaction_at']]

    # Dividends
    dividends = robin_stocks.robinhood.get_dividends()
    dividends = pd.DataFrame.from_records(dividends)
    dividends[['Name', 'Simple Name', 'Symbol']] = dividends.apply(
        lambda x: get_robinhood_instrument(x['instrument']), result_type='expand', axis=1)

    dividends = dividends[['Name', 'Simple Name',
                           'Symbol', 'state', 'amount', 'payable_date']]

    robin_stocks.robinhood.authentication.logout()

    return option_orders, dividends


def get_robinhood_instrument(url: str) -> tuple[str, str, str]:
    """Gets the Stock Information for each transaction

    :param url: Instrument URL
    :type url: str
    :return: Table consisting of name, simple_name and symbol for the url
    :rtype: tuple[str, str, str]
    """
    instrument = robin_stocks.robinhood.get_instrument_by_url(url)

    return instrument['name'], instrument['simple_name'], instrument['symbol']
