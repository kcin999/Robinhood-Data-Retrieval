import os
import shutil
import datetime

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

	shutil.copy(root + "\\" +config.get_csv_fileName(),root + "\\archive\\"+ datetimeNow.strftime("%Y") + " " +datetimeNow.strftime("%b") +"\\CSV_backup "+ datetimeNow.strftime("%H") + datetimeNow.strftime("%M") + datetimeNow.strftime("%S") + ".csv")


def createArchiveFolder():
	datetimeNow = datetime.datetime.now()

	path = getCurrentDir() + r"\archive"

	createFolder(path)

	path = getCurrentDir() + r"\archive" + "\\" + datetimeNow.strftime("%Y") + " " +datetimeNow.strftime("%b")

	createFolder(path)

def archiveAll():
	createArchiveFolder()
	archiveCSV()