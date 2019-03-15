# coding: utf-8
'''
Author: 	   Ren
Created Date:  Apr. 13, 2017
Modified Date: Aug. 29, 2017

Specification:
Core function. Apply statistics on and make excels and pdfs from txt files. 
The txt files to be processed should be gather in a folder, which's path should be given below to "path".
'''

import re
import os
import json
import xlsxwriter
import pypandoc
import glob
from pprint import pprint

path = 'dataset/20171127_1/' # path of the folder that contains txt files that should be processed.

# -------------------------------------------- latex text -------------------------------------------
latexHead = r'''
\documentclass{article}
\usepackage[UTF8,noindent]{ctex}
\usepackage{geometry}
\geometry{
 a4paper,
 total={170mm,257mm},
 left=20mm,
 top=30mm,
 bottom=30mm	
 }
\usepackage[usenames, dvipsnames]{color}
\definecolor{mypink1}{rgb}{0.858, 0.188, 0.478}
\definecolor{mypink2}{RGB}{219, 48, 123}
\definecolor{mypink3}{cmyk}{0, 0.7808, 0.4429, 0.1412}
\definecolor{mygray}{gray}{0.6}
\usepackage{fontspec}
\setmainfont{Arial} 
\setsansfont{Arial} 
\setmonofont{Arial} 
\setCJKmainfont{Hiragino Sans GB} 
\setCJKsansfont{Hiragino Sans GB} 
\setCJKmonofont{Hiragino Sans GB} 

\setlength{\parindent}{0pt}
\setlength{\parsep}{1\baselineskip}
\begin{document}

'''
latexEnd = r'''
\end{document}
'''


# ---------------------------------------- file specification ---------------------------------------
inputFileNames = glob.glob(path + "*.txt")


# ---------------------------------------- load reference file --------------------------------------
kywdDict = {}
kywdRef  = {}
kywdList = []

with open("key_word_dict_extended/structured_dict.json") as jsonFile:
	kywdDict = json.load(jsonFile)

with open("key_word_dict_extended/reference_dict.json") as jsonFile:
	kywdRef  = json.load(jsonFile)

with open("key_word_dict_extended/stop_word_dict.json") as jsonFile:
	stopDict = json.load(jsonFile)

# Sorted list of keywords. Sort according to the length of keyword. The longer the keyword is, the earlier the keyword appears.
kywdList = sorted(kywdRef.keys(), key=len, reverse=True)

# ------------------------- directly statistic & generate latex file --------------------------------
for inputFileName in inputFileNames:
	print("Processing", inputFileName, "...")
	inputFileName = inputFileName.split(".")[0]
	isIni = True
	paragraphIndex = 0
	kywdStat = {}
	outputFileName = inputFileName + '_processed'
	with open(inputFileName + ".txt") as curFile:
		with open(outputFileName + ".tex", "w") as outputFile:
			outputFile.write(latexHead)
			for curParagraph in curFile:
				curParagraph = re.sub(r"%", r"\\%", curParagraph)
				curParagraph = re.sub(r"\:", r"\\\:", curParagraph)
				curParagraph = re.sub(r"\^", "\^", curParagraph)
				curParagraph = re.sub(r"_", r"\\_", curParagraph)
				curParagraph = re.sub(r"#", r"\\#", curParagraph)

				# Jump off the empty paragraph.
				if re.match(r"\s*\n", curParagraph):
					continue

				# Default: the first paragraph of a file is the title.
				# if isIni:
				# 	isIni = False
				# 	curParagraph = re.sub(r"^\s*", r"{\\Huge \\bfseries ", curParagraph)
				# 	curParagraph = re.sub(r"\n$", r"}\\\\ \\\\\n", curParagraph)
				# 	outputFile.write(curParagraph)
				# 	continue

				paragraphIndex += 1
				paragraphName = "Para " + str(paragraphIndex)
				kywdStat[paragraphName] = {
					"found_keyword": False,
					"related_industry": [],
					"first_keyword_position": {},
					"statistic_result": {},
					"statistic_code": {}
				}

				# Set beginning and ending of a paragraph. Add para tag. Add an extra line after this paragraph.
				curParagraph = re.sub(r"^\s*", r"\\textcolor{mygray}{\\textit{< " + paragraphName + r" >}}\\\\", curParagraph)
				curParagraph = re.sub(r"\n$", r"\\\\ \\\\\n", curParagraph)

				# Statistic of each keyword
				for keyword in kywdList:
					# if the keyword is in the stop words list, just skip this word, no matter what the detail code is.
					if keyword in stopDict:
						continue

					codeList = set(kywdRef[keyword])

					# if the keyword is in the stop words list, check the stop code and delete them from the ref code list.
					# if keyword in stopDict:
					# 	stopList = set(stopDict[keyword])
					# 	codeList = codeList - stopList

					codeList = list(codeList)

					if re.search(r'(?<!z)' + keyword + r'(?!m)', curParagraph):
						kywdStat[paragraphName]["found_keyword"] = True
						kywdStat[paragraphName]["statistic_result"][keyword] = len(re.findall(r'(?<!z)' + keyword + r'(?!m)', curParagraph))
						curParagraph = re.sub(r'(?<!z)' + keyword + r'(?!m)', r"\\textcolor{red}{\\bfseries " + "mdzz".join(keyword) + "(" + ",".join(codeList) + ") }", curParagraph)

				for keyword in kywdStat[paragraphName]["statistic_result"].keys():
					kywdStat[paragraphName]["first_keyword_position"][keyword] = re.search(r'(?<!z)' + "mdzz".join(keyword) + r'(?!m)', curParagraph).span()[0]

				# print("curParagraph: ", curParagraph)

				curParagraph = re.sub("mdzz", "", curParagraph)
				outputFile.write(curParagraph)

			outputFile.write(latexEnd)

	os.system("xelatex -interaction=batchmode -output-directory=" + path + " " + outputFileName + ".tex")



	# -------------------------------------- analysis statistic -----------------------------------------
	def get_industry_from_index (indexs):
		outputIndustries = []
		for index in indexs:
			if index[0:2] not in outputIndustries:
				outputIndustries.append(index[0:2])
		return outputIndustries

	# pprint(kywdStat)

	for paragraphName in kywdStat:
		if kywdStat[paragraphName]['found_keyword']:
			statResult = kywdStat[paragraphName]['statistic_result']
			keywordList = sorted(kywdStat[paragraphName]['first_keyword_position'].items(), key = lambda x: x[1])
			for keywordSet in keywordList:
				keyword = keywordSet[0]
				for industryIndex in get_industry_from_index(kywdRef[keyword]):
					if industryIndex not in kywdStat[paragraphName]['related_industry']:
						kywdStat[paragraphName]['related_industry'].append(industryIndex)
				for codeIndex in kywdRef[keyword]:
					if codeIndex not in kywdStat[paragraphName]['statistic_code']:
						kywdStat[paragraphName]['statistic_code'][codeIndex] = kywdStat[paragraphName]['statistic_result'][keyword]
					else:
						kywdStat[paragraphName]['statistic_code'][codeIndex] += kywdStat[paragraphName]['statistic_result'][keyword]

	# with open("keywords_statistic.json", "w") as jsonFile:
	# 	jsonFile.write(json.dumps(kywdStat, indent = 4, ensure_ascii = False))


	# -------------------------------------- generate excel file ----------------------------------------
	workbook = xlsxwriter.Workbook(outputFileName + ".xlsx")
	worksheet = workbook.add_worksheet()

	worksheet.set_column('A:D', 7)
	worksheet.set_column('E:E', 20)
	worksheet.set_column('F:F', 7)
	worksheet.set_column('G:G', 10)
	worksheet.set_column('H:W', 4)
	worksheet.set_column('K:K', 7.5)
	worksheet.set_column('P:P', 7)
	worksheet.set_column('T:T', 7.5)

	title_format = workbook.add_format({
		'font_name': "DengXian",
		'bold': 1,
		'align': 'center',
		'valign': 'vcenter',
		'num_format': '@'
		})
	title_with_bottom_format = workbook.add_format({
		'font_name': "DengXian",
		'bold': 1,
		'align': 'center',
		'valign': 'vcenter',
		'bottom': 1,
		'bottom_color': '#555555',
		'num_format': '@'
		})
	title_with_right_format = workbook.add_format({
		'font_name': "DengXian",
		'bold': 1,
		'align': 'center',
		'valign': 'top',
		'right': 1,
		'right_color': '#555555',
		'num_format': '@',
		'bottom': 2,
		'bottom_color': '#333333'
		})
	content_format = workbook.add_format({
		'font_name': "DengXian",
		'valign': 'vcenter',
		'num_format': '@'
		})
	content_with_bottom_format = workbook.add_format({
		'font_name': "DengXian",
		'valign': 'vcenter',
		'num_format': '@',
		'bottom': 2,
		'bottom_color': '#333333'
		})
	keyword_format = workbook.add_format({
		'font_name': "DengXian",
		'valign': 'vcenter',
		'bold': 1,
		'font_color': 'red',
		'num_format': '@'
		})
	keyword_with_bottom_format = workbook.add_format({
		'font_name': "DengXian",
		'valign': 'vcenter',
		'bold': 1,
		'font_color': 'red',
		'num_format': '@',
		'bottom': 2,
		'bottom_color': '#333333'
		})
	title_format_with_yellow_bgd = workbook.add_format({
		'font_name': "DengXian",
		'bold': 1,
		'align': 'center',
		'valign': 'vcenter',
		'bg_color': '#FFFD78',
		'border': 1,
		'border_color': '#CCCCCC',
		'num_format': '@'
		})
	title_with_bottom_format_with_yellow_bgd = workbook.add_format({
		'font_name': "DengXian",
		'bold': 1,
		'align': 'center',
		'valign': 'vcenter',
		'border': 1,
		'border_color': '#CCCCCC',
		'bottom': 1,
		'bg_color': '#FFFD78',
		'bottom_color': '#555555',
		'num_format': '@'
		})

	worksheet.merge_range("A1:A2", "段落标号", title_with_bottom_format)
	worksheet.merge_range('B1:E1', "涉及行业", title_format)
	worksheet.merge_range("F1:F2", "出现次数", title_with_bottom_format)
	worksheet.merge_range("G1:G2", "其他关键词（以'；'分隔）", title_with_bottom_format)
	worksheet.merge_range("H1:N1", "具体举措（计数）", title_format)
	worksheet.merge_range("O1:O2", "位置", title_with_bottom_format)
	worksheet.merge_range("P1:W1", "负面态度（计数）", title_format_with_yellow_bgd)
	worksheet.write("B2", "2位代码", title_with_bottom_format)
	worksheet.write("C2", "3位代码", title_with_bottom_format)
	worksheet.write("D2", "4位代码", title_with_bottom_format)
	worksheet.write("E2", "行业名称", title_with_bottom_format)
	worksheet.write("H2", "园区", title_with_bottom_format)
	worksheet.write("I2", "项目", title_with_bottom_format)
	worksheet.write("J2", "企业", title_with_bottom_format)
	worksheet.write("K2", "产品/服务", title_with_bottom_format)
	worksheet.write("L2", "技术", title_with_bottom_format)
	worksheet.write("M2", "其他", title_with_bottom_format)
	worksheet.write("N2", "待定", title_with_bottom_format)
	worksheet.write("P2", "泛指次数", title_with_bottom_format_with_yellow_bgd)
	worksheet.write("Q2", "园区", title_with_bottom_format_with_yellow_bgd)
	worksheet.write("R2", "项目", title_with_bottom_format_with_yellow_bgd)
	worksheet.write("S2", "企业", title_with_bottom_format_with_yellow_bgd)
	worksheet.write("T2", "产品/服务", title_with_bottom_format_with_yellow_bgd)
	worksheet.write("U2", "技术", title_with_bottom_format_with_yellow_bgd)
	worksheet.write("V2", "其他", title_with_bottom_format_with_yellow_bgd)
	worksheet.write("W2", "待定", title_with_bottom_format_with_yellow_bgd)

	curExcelRow = 2
	curExcelContent = {}
	curExcelStyle = {}

	for paragraph in kywdStat:
		if not kywdStat[paragraph]["found_keyword"]:
			worksheet.write(curExcelRow, 0, paragraph, title_with_right_format)
			for i in range(1, 23):
				worksheet.write(curExcelRow, i, "", content_with_bottom_format)
			curExcelRow += 1
		else:
			startExcelRow = curExcelRow
			curIndustryIndexs = kywdStat[paragraph]["related_industry"]
			curFoundKeycodes  = kywdStat[paragraph]["statistic_code"]
			for industryIndex in curIndustryIndexs:
				curExcelContent = {}
				curExcelStyle = {}
				fstIndustry = kywdDict[industryIndex]

				worksheet.write(curExcelRow, 1, fstIndustry["index"], content_format)
				curExcelContent[1] = fstIndustry["index"]
				if fstIndustry["index"] in curFoundKeycodes:
					worksheet.write(curExcelRow, 2, "", content_format)
					worksheet.write(curExcelRow, 3, "", content_format)
					worksheet.write(curExcelRow, 4, fstIndustry["name"], keyword_format)
					curExcelContent[4] = fstIndustry["name"]
					curExcelStyle[4] = "keyword_format"
					worksheet.write(curExcelRow, 5, curFoundKeycodes[fstIndustry["index"]], content_format)
					curExcelContent[5] = curFoundKeycodes[fstIndustry["index"]]
				else:
					worksheet.write(curExcelRow, 2, "", content_format)
					worksheet.write(curExcelRow, 3, "", content_format)
					worksheet.write(curExcelRow, 4, fstIndustry["name"], content_format)
					curExcelContent[4] = fstIndustry["name"]
					worksheet.write(curExcelRow, 5, "", content_format)
				for i in range(6, 23):
					worksheet.write(curExcelRow, i, "", content_format)

				curExcelRow += 1
				secIndustryIndexs = fstIndustry["child"]
				for secIndustryIndex in secIndustryIndexs:
					curExcelContent = {}
					curExcelStyle = {}
					secIndustry = secIndustryIndexs[secIndustryIndex]

					worksheet.write(curExcelRow, 1, "", content_format)
					worksheet.write(curExcelRow, 2, secIndustry["index"], content_format)
					curExcelContent[2] = secIndustry["index"]
					if secIndustry["index"] in curFoundKeycodes:
						worksheet.write(curExcelRow, 3, "", content_format)
						worksheet.write(curExcelRow, 4, secIndustry["name"], keyword_format)
						curExcelContent[4] = secIndustry["name"]
						curExcelStyle[4] = "keyword_format"
						worksheet.write(curExcelRow, 5, curFoundKeycodes[secIndustry["index"]], content_format)
						curExcelContent[5] = curFoundKeycodes[secIndustry["index"]]
					else:
						worksheet.write(curExcelRow, 3, "", content_format)
						worksheet.write(curExcelRow, 4, secIndustry["name"], content_format)
						curExcelContent[4] = secIndustry["name"]
						worksheet.write(curExcelRow, 5, "", content_format)
					for i in range(6, 23):
						worksheet.write(curExcelRow, i, "", content_format)

					trdIndustryIndexs = secIndustry["child"]
					if len(trdIndustryIndexs) < 2:
						worksheet.write(curExcelRow, 3, secIndustry["index"] + "0", content_format)
						curExcelContent[3] = secIndustry["index"] + "0"
						curExcelRow += 1
					else:
						curExcelRow += 1
						for trdIndustryIndex in trdIndustryIndexs:
							curExcelContent = {}
							curExcelStyle = {}
							trdIndustry = trdIndustryIndexs[trdIndustryIndex]

							worksheet.write(curExcelRow, 1, "", content_format)
							worksheet.write(curExcelRow, 2, "", content_format)
							worksheet.write(curExcelRow, 3, trdIndustry["index"], content_format)
							curExcelContent[3] = trdIndustry["index"]
							if trdIndustry["index"] in curFoundKeycodes:
								worksheet.write(curExcelRow, 4, trdIndustry["name"], keyword_format)
								curExcelContent[4] = trdIndustry["name"]
								curExcelStyle[4] = "keyword_format"
								worksheet.write(curExcelRow, 5, curFoundKeycodes[trdIndustry["index"]], content_format)
								curExcelContent[5] = curFoundKeycodes[trdIndustry["index"]]
							else:
								worksheet.write(curExcelRow, 4, trdIndustry["name"], content_format)
								curExcelContent[4] = trdIndustry["name"]
								worksheet.write(curExcelRow, 5, "", content_format)
							for i in range(6, 23):
								worksheet.write(curExcelRow, i, "", content_format)

							curExcelRow += 1

			if not startExcelRow + 1 == curExcelRow:
				worksheet.merge_range("A" + str(startExcelRow + 1) + ":A" + str(curExcelRow), paragraph, title_with_right_format)
			else:
				worksheet.write(startExcelRow, 0, paragraph, title_with_right_format)
			for i in range(1, 23):
				if i in curExcelContent:
					if i in curExcelStyle:
						worksheet.write(curExcelRow - 1, i, curExcelContent[i], keyword_with_bottom_format)
					else:
						worksheet.write(curExcelRow - 1, i, curExcelContent[i], content_with_bottom_format)
				else:
					worksheet.write(curExcelRow - 1, i, "", content_with_bottom_format)



	worksheet.freeze_panes(2, 1)

	workbook.close()

