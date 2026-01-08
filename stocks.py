"""Stock Functions"""
# Module Imports
import os
from datetime import datetime, timedelta
import robin_stocks
import pyotp
import yfinance
from alpha_vantage.timeseries import TimeSeries
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


def get_historical_data(ticker: str, startdate: datetime = None, enddate: datetime = None,  interval: str = '5m') -> pd.DataFrame:
    """Gets historical data for a given stock ticker.

    :param ticker: Ticker to get data for
    :type ticker: str
    :param startdate: Start Date of range to get data for, defaults to None

        * If startdate and enddate are null, data will be returned over the last 7 days
        * If startdate is given and enddate is null, 7 days of data will be returned starting from the start date

    :type startdate: datetime, optional
    :param enddate: Start Date of range to get data for, defaults to None
    :type enddate: datetime, optional
    :param interval: Time between records, defaults to '5m'
    :type interval: str, optional
    :return: _description_
    :rtype: pd.DataFrame
    """
    # If both startdate and enddate are null, get data over the last 7 days
    if startdate is None and enddate is None:
        startdate = datetime.now() - timedelta(days=7)
        enddate = datetime.now()

    # If startdate and not enddate, get 7 days worth of data from the startdate
    if startdate and enddate is None:
        enddate = startdate + timedelta(days=7)

    #  If startdate is not given and enddate is, get 7 days worth of data
    if startdate is None and enddate:
        startdate = enddate - timedelta(days=7)

    # 5 minute interval checks
    if interval == '5m' and (datetime.now() - timedelta(days=60)) > startdate:
        raise ValueError("For a 5 minute interval, startdate must be within the last 60 days")

    # 1 minute interval checks
    if interval == '1m':
        if (datetime.now() - timedelta(days=30)) > startdate:
            raise ValueError("For a 1 minute interval, startdate must be within the last 30 days")
        
        if (enddate - startdate).days > 7:
            raise ValueError("For a 1 minute interval, only 7 days of data can be grabbed at a time")
        
    data = yfinance.download(tickers=ticker, start=startdate, end=enddate, interval=interval)

    return data


def get_robinhood_transactions() -> tuple[pd.DataFrame, pd.DataFrame]:
    """Returns Robinhood Transactions from the user account

    :return: Transactions and Divdend dataframes
    :rtype: tuple[pd.DataFrame, pd.DataFrame]
    """
    totp = pyotp.TOTP(os.environ.get('ROBINHOOD_ONETIME_PASSWORD')).now()
    _login = robin_stocks.robinhood.authentication.login(
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

