# coding: utf-8

import re
import os
import glob
import json
import xlsxwriter
import pypandoc
import subprocess
from pprint import pprint


path = '../dataset/20170627-rerun/'
path_done = '../dataset/20170627/'

# tex_set = set([])
# pdf_set = set([])

# for fileName in os.listdir(path):
# 	if re.match(r'.*\.tex', fileName):
# 		tex_set.add(fileName)
# 	elif re.match(r'.*\.pdf', fileName):
# 		pdf_set.add(fileName)
# 	else:
# 		pass

# for fileName in tex_set:
# 	if not (fileName.split(".")[0] + ".pdf") in pdf_set:
# 		os.system("mv " + path + fileName + " " + path_rerun + fileName)

tex_set = set([])
txt_set = set([])

for fileName in os.listdir(path):
	if re.match(r'.*\.tex', fileName):
		tex_set.add(fileName)
	elif re.match(r'.*\.txt', fileName):
		txt_set.add(fileName)
	else:
		pass

for fileName in txt_set:
	if not (fileName.split(".")[0] + "_processed.tex") in tex_set:
		os.system("mv " + path + fileName + " " + path_done + fileName)
