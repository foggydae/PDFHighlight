# coding: utf-8
'''
Author: 	   Ren
Created Date:  Jun. 26, 2017
Modified Date: Aug. 29, 2017

Specification:
Convert UTF-16 encoding type into UTF-8.
'''

import re
import os
import glob
import json
import xlsxwriter
import pypandoc
import subprocess
from pprint import pprint



path = '../dataset/20170727/'

for fileName in os.listdir(path):
	if not re.match(r'.*\.txt', fileName):
		print("continue")
		continue

	newfileName = path + "./new/" + fileName.split(".")[0] + ".txt"

	fileName = path + fileName

	filetype = subprocess.check_output("file -I " + fileName + " |awk -F '=' '{print $2}'", shell=True)
	print(filetype)
	if filetype == b'utf-16le\n':
		os.system("iconv -c -f utf-16le -t UTF-8 " + fileName + " > " + newfileName)
	else:
		pass