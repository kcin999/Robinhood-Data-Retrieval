#Module Imports
import os
import shutil
import datetime

#File Imports
import config

def getCurrentDir():
	return os.path.dirname(os.path.abspath(__file__))

def createFolder(path):
	if os.path.isdir(path) == False:
		try:
			os.mkdir(path)

		except OSError:
			print("Creation of the directory %s failed" % path)
		else:
			print("Successfully created the directory %s" % path)

def archiveCSV():
	root = getCurrentDir()
	datetimeNow = datetime.datetime.now()

	pathToFolder = root + "\\archive\\"+ datetimeNow.strftime("%Y") + " " +datetimeNow.strftime("%b") + "\\" + datetimeNow.strftime("%Y %b %d")

	shutil.copy(root + "\\" +config.get_csv_myStocksFileName(),pathToFolder + "\\Backup_" + config.get_csv_myStocksFileName() + datetimeNow.strftime("%H") + datetimeNow.strftime("%M") + datetimeNow.strftime("%S") + ".csv")
	shutil.copy(root + "\\" +config.get_csv_marketStockDataFileName(),pathToFolder + "\\Backup_" + config.get_csv_marketStockDataFileName() + datetimeNow.strftime("%H") + datetimeNow.strftime("%M") + datetimeNow.strftime("%S") + ".csv")

def createArchiveFolder():
	datetimeNow = datetime.datetime.now()

	path = getCurrentDir() + r"\archive"

	createFolder(path)

	path = getCurrentDir() + r"\archive" + "\\" + datetimeNow.strftime("%Y") + " " +datetimeNow.strftime("%b")

	createFolder(path)

	path = path + "\\" + datetimeNow.strftime("%Y %b %d")
	createFolder(path)

def archiveAll():
	createArchiveFolder()
	archiveCSV()