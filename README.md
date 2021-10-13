# Robinhood-Data-Retrieval
This is a repository to manage my stock spreadsheet.  I wanted to get it all automated, so that I would not have to update it each and every day manually.  

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

Can be installed using
```
pip install robin_stocks
pip install toml
```
## Notes:
* You can also run [createNewSheet.py] to get started.  Given credentials, it will login to your Robinhood, and create a new sheet, named the current year for the current stocks that you have.
* If you buy a new stock, you **MUST** add it to spreadsheet.
	* Right Click - Insert

## To run daily update:
1. Make sure libraries are installed.
2. Make sure the excel file is in the same directory as the code.
	* It can be named whatever you would like, the name just has to be in the config.toml.
3. Update the config.toml file contain the correct information.
4. Run [main.py](main.py)

## To set up for the first time:
1. Make sure python and libraries are installed. 
2. Update the config.toml file contain the correct information. 
	* Information can be found on how to get your mfa_code here [Robin_Stocks Github] (https://github.com/jmfernandes/robin_stocks/blob/master/Robinhood.rst#with-mfa-entered-programmatically-from-time-based-one-time-password-totp)
3. Run [createNewSheet.py](createNewSheet.py)


### Written by:
[Nick Bierman](https://github.com/kcin999)
