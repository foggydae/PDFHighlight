# coding: utf-8

'''
Author: 	   Ren
Created Date:  Jul. 29, 2017
Modified Date: Aug. 29, 2017

Specification:
Partially solve the paragraphing issue occurred on those transformed from Word -> Work in most cases.
'''

import re
import os
from pprint import pprint

sourcePath = "../dataset/20171127_1/"
targetPath = "../dataset/transformed/"

for fileName in os.listdir(sourcePath):
	if re.match(r'.*\.txt', fileName):
		with open(sourcePath + fileName) as curFile:
			with open(targetPath + fileName, "w") as outputFile:
				for curParagraph in curFile:
					# curParagraph = re.sub(r"\s", "mdzz", curParagraph)
					curParagraph = re.sub(r"\n", "mdzz", curParagraph)
					curParagraph = re.sub(" ", "mdzz", curParagraph)
					curParagraph = re.sub("　", "mdzz", curParagraph)
					curParagraph = re.sub("  ", "mdzz", curParagraph)
					curParagraph = re.sub(r"[mdzz]+", "\n\n", curParagraph)
					outputFile.write(curParagraph)






