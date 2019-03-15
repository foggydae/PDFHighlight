# coding: utf-8
'''
Author: 	   Ren
Created Date:  Jun. 27, 2017
Modified Date: Aug. 29, 2017

Specification:
Using commend 'textutil' to convert word files into txt. 
'''

import re
import os
import glob
import json
import xlsxwriter
import pypandoc
import subprocess
from pprint import pprint



path = '../dataset/20171127/'
path_done = '../dataset/20171127/'
# path_done = '../word/'

for fileName in os.listdir(path):
	if re.match(r'.*\.doc', fileName):
		os.system("textutil -convert txt " + path + fileName)
	else:
		pass

for fileName in os.listdir(path):
	if re.match(r'.*\.txt', fileName):
		with open(path + fileName) as curFile:
			with open(path_done + fileName, "w") as outputFile:
				for curParagraph in curFile:
					curParagraph = re.sub(r"　　", "\n", curParagraph)
					curParagraph = re.sub(r"     ", "\n", curParagraph)
					curParagraph = re.sub(r"      ", "\n", curParagraph)
					outputFile.write(curParagraph)






