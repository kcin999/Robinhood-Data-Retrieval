#File Imports
import stocks
import excel

def main():
	stockList = stocks.getRobinhoodStocks()
	excel.createSheet(stockList)

if __name__ == "__main__":
	main()