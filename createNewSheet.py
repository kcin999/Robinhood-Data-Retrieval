#File Imports
import robinhood
import excel


def main():
	stocks = robinhood.getStocks()
	excel.createSheet(stocks)

if __name__ == "__main__":
	main()