# coding: utf-8

'''
Author: 	   Ren
Created Date:  Jul. 29, 2017
Modified Date: Aug. 29, 2017

Specification:
Gather txt files with to-be-solved issues. the files' number(name) are given within a given excel.
'''

import re
import xlrd
import os
from pprint import pprint

excelFileName = "../left-issues.xlsx"
datasetPath = "../dataset/"
datasetFolders = ["20170829", "20170828", "20170812", "20170731", "20170727", "20170729", "20170630", "20170628", "20170715", "20170627_rerun", "20170627", "20170623", "20170621", "20170531", "20170526", "20170517", "20170505"]

kywdFile = xlrd.open_workbook(excelFileName)
for i in range(0, 1): # insruction: the second number(now is 1) indicate how many sheets in the excel will be processed. 
	kywdTable = kywdFile.sheets()[i]
	kywdRows = kywdTable.nrows
	kywdCols = kywdTable.ncols
	folderName = kywdTable.name
	print(folderName)

	targetPath = "../dataset/" + folderName + "/"
	os.system("mkdir " + folderName)

	for rowIndex in range(1, kywdRows):
		rowValue = kywdTable.row_values(rowIndex)
		fileName = rowValue[0] + ".txt"
		searchFlag = 0

		for datasetFolder in datasetFolders:
			curSearchPath = datasetPath + datasetFolder + "/"
			fileset = os.listdir(curSearchPath)
			if fileName in fileset:
				os.system("cp " + curSearchPath + fileName + " " + targetPath + fileName)
				searchFlag = 1
				break

		if searchFlag == 0:
			print(fileName + " fail to search")