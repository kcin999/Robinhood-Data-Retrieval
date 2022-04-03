#Module Imports
import datetime
import logging
import sys
import traceback
#File Imports
import excel
import stocks
import config
import archive
import updateCSV

def main():
	if datetime.datetime.now().weekday() != 6 and datetime.datetime.now().weekday() != 5:
		myStocks = stocks.getRobinhoodStocks()
		excel.addDataToExcel(myStocks)
		updateCSV.writeToCSV(myStocks)
		archive.archiveAll()


if __name__ == "__main__":
	try:
		main()
	except Exception as e:
		exc_type, exc_obj, exc_tb = sys.exc_info()
		logging.info("\t\t" + str(e) + " at " + datetime.datetime.now().strftime("%m/%d/%Y %H:%M:%S:%f"))
		logging.info("Call Stack:")
		callStack = traceback.format_exc()
		logging.info( callStack + "\n\n")