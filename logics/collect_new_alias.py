# coding = utf-8
'''
Author: 	   Ren
Created Date:  May. 04, 2017
Modified Date: Aug. 29, 2017

Specification:
Collect new alias from excels handed in by RAs, and merge them into a "result_stat.csv" file.
Then the .csv file can be convert into an .xlsx file manually, using the Excel.
'''

import re
import xlrd
import os
import json
import csv
import xlsxwriter
from pprint import pprint

global RESULT_FILE_LIST
RESULT_FILE_LIST = []
RESULT_FILES_PATH = '../dataset/20170426/Result' # path of the folder that contains all the excel files handed in by RAs


def gci(filepath):
	global RESULT_FILE_LIST
	files = os.listdir(filepath)
	for fi in files:
		fi_d = os.path.join(filepath,fi)            
		if os.path.isdir(fi_d):
			gci(fi_d)                  
		else:
			if fi_d.split('.').pop() == 'xlsx':
				RESULT_FILE_LIST.append(fi_d)

gci(RESULT_FILES_PATH)

newAliasList = []

for fileName in RESULT_FILE_LIST:
	print(fileName)

	kywdFile = xlrd.open_workbook(fileName)
	kywdTable = kywdFile.sheets()[0]
	kywdRows = kywdTable.nrows
	kywdCols = kywdTable.ncols

	for rowIndex in range(1, kywdRows):
		rowValue = kywdTable.row_values(rowIndex)
		if rowValue[6] ==  "":
			continue
		
		newAlias = rowValue[6].split('；')
		accCode = '-1'
		for colIndex in range(3, 0, -1):
			if not rowValue[colIndex] == "":
				accCode = rowValue[colIndex]
		if accCode == '-1':
			print("ERROR. File '", fileName, "' has new alias without code at row #", rowIndex, ". Please check the format.")
		else:
			accName = rowValue[4]
			for newAlia in newAlias:
				newAliasList.append([newAlia, accCode, accName, fileName])

# newAliasList = list(set(newAliasList))
pprint(newAliasList)

with open(RESULT_FILES_PATH + "/result_stat.csv", 'w') as csvFile:
	writer = csv.writer(csvFile)
	writer.writerow(['New Alias', 'Code', 'Name'])
	writer.writerows(newAliasList)
