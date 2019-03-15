# coding: utf-8
'''
Author: 	   Ren
Created Date:  May. 30, 2017
Modified Date: Aug. 29, 2017

Specification:
Ignore txt files which is before 2000.
'''

import re
import os
import csv
import glob
import json
import xlsxwriter
import pypandoc
import subprocess
from pprint import pprint


path = '../dataset/20170630/'

for fileName in os.listdir(path):
	if not re.match(r'.*\.txt', fileName):
		print("continue")
		continue

	fileNameStr = fileName.split(".")[0]
	year = int(fileNameStr.split("_")[1])

	if year < 2000:
		os.system("mv " + path + fileName + " ../old-ignore/" + fileName)
		os.system("mv " + path + fileName.split(".")[0] + "_processed.xlsx" + " ../old-ignore/" + fileName.split(".")[0] + "_processed.xlsx")
		os.system("mv " + path + fileName.split(".")[0] + "_processed.pdf" + " ../old-ignore/" + fileName.split(".")[0] + "_processed.pdf")
