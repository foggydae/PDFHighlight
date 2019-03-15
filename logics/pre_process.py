# coding: utf-8
'''
Author: 	   Ren
Created Date:  Apr. 24, 2017
Modified Date: Aug. 29, 2017

Specification:
Convert different encoding type into UTF-8. Work for most files.
'''

import re
import os
import glob
import json
import xlsxwriter
import pypandoc
import subprocess
from pprint import pprint



path = '../dataset/20171127_2/'

for fileName in os.listdir(path):
	# print(fileName)
	if not re.match(r'.*\.txt', fileName):
		print("continue")
		continue

	outfileName = path + "transformed/" + fileName.split(".")[0] + ".txt"
	errfileName = path + "error/" + fileName.split(".")[0] + ".txt"
	fileName = path + fileName

	checkRTF = subprocess.check_output("file -bI " + fileName + "|awk -F ';' '{print $1}'", shell=True)
	# os.system("file -bI " + fileName + "|awk -F ';' '{print $1}'")
	if checkRTF == b'text/rtf\n':
		os.system("textutil -convert txt " + fileName)

	filetype = subprocess.check_output("file -I " + fileName + " |awk -F '=' '{print $2}'", shell=True)
	# filetype = os.system("file -I " + fileName + " |awk -F '=' '{print $2}'")

	if filetype == b'utf-8\n':
		os.system("cp " + fileName + " " + outfileName)
	elif filetype == b'iso-8859-1\n':
		os.system("iconv -c -f GB2312 -t UTF-8 " + fileName + " > " + outfileName)
	elif filetype == b'unknown-8bit\n':
		os.system("iconv -c -f GB2312 -t UTF-8 " + fileName + " > " + outfileName)		
	else:
		os.system("cp " + fileName + " " + errfileName)
		print(filetype)