# Robinhood-Data-Retrieval
This is a repository to manage my stock spreadsheet.  I wanted to get it all automated, so that I would not have to update it each and every day manually.  

Also using Microsoft PowerBI now. When the [main.py](main.py) is ran, it not only updated the excel spreadhheet

## Libraries used
* robin_stocks
	* Documentation can be found here: 
		1. [Robin Stocks Github](https://github.com/jmfernandes/robin_stocks) 
		2. [Robin-Stocks.com](http://www.robin-stocks.com/en/latest/robinhood.html)
* toml
	* Documentation can be found here:
		1. [TOML Github](https://github.com/toml-lang/toml)
* openpyxl
	* Documentation can be found here:
		1. [readthedocs.io](https://openpyxl.readthedocs.io/en/stable/)
* yfinance

Can be installed using pip. See below:
```
pip install robin_stocks
pip install toml
pip install yfinance
```

## To run daily update:
1. Make sure libraries are installed.
2. Make sure the excel file is in the same directory as the code.
	* It can be named whatever you would like, the name just has to be in the config.toml.
3. Update the config.toml file contain the correct information.
4. Run [main.py](main.py)
	* This can be ran using a task scheduler

## To set up for the first time:
1. Make sure python and libraries are installed. 
2. Update the config.toml file contain the correct information. 
	* Information can be found on how to get your mfa_code here [Robin_Stocks Github] (https://github.com/jmfernandes/robin_stocks/blob/master/Robinhood.rst#with-mfa-entered-programmatically-from-time-based-one-time-password-totp)
3. Running [main.py](main.py) will create the Excel Document as well as the Sheet if it does not exist already.
	* Adds stocks that already existing in your robinhood account
	* On the start of a new year, a new sheet will also be created

## Features to come:
1. Calculate Percentage of Portfolio that each stock takes up.

### Written by:
[Nick Bierman](https://github.com/kcin999)