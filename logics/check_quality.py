# coding: utf-8
'''
Author: 	   Ren
Created Date:  May. 30, 2017
Modified Date: Aug. 29, 2017

Specification:
Check the amount of paragraph a txt has to indicate if the document has a paragraphing issue
'''

import re
import os
import glob
import json
import xlsxwriter
import pypandoc
import subprocess
from pprint import pprint



path = '../dataset/20170715/'

for fileName in os.listdir(path):
	print(fileName)
	if not re.match(r'.*\.txt', fileName):
		print("continue")
		continue

	errfileName = path + "../error/new/" + fileName.split(".")[0] + ".txt"
	checkfileName = path + "../2Bcheck/" + fileName.split(".")[0] + ".txt"
	fileName = path + fileName

	count = 0
	with open(fileName) as curFile:
		for curParagraph in curFile:
			count += 1
	print(count)

	if count < 10:
		os.system("mv " + fileName + " " + errfileName)
	elif count > 500:
		os.system("mv " + fileName + " " + errfileName)
	elif count > 250:
		os.system("mv " + fileName + " " + checkfileName)
	else:
		pass

		