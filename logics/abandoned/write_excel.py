# coding: utf-8

import re
import os
import json
import xlsxwriter
import pypandoc
from pprint import pprint

with open("key_word_dict/industry_dict.json") as jsonFile:
	kywdDict = json.load(jsonFile)

with open("keywords_statistic.json") as jsonFile:
	kywdStat = json.load(jsonFile)
pprint(kywdStat)

inputFileName = 'sample/2005年唐山人民政府工作报告'
outputFileName = inputFileName + '_out'



workbook = xlsxwriter.Workbook(outputFileName + ".xlsx")
worksheet = workbook.add_worksheet()

worksheet.set_column('A:Z', 10)
worksheet.set_column('E:E', 30)
worksheet.set_column('G:G', 30)

title_format = workbook.add_format({
	'font_name': "DengXian",
	'bold': 1,
	'align': 'center',
	'valign': 'vcenter'
	})
title_with_bottom_format = workbook.add_format({
	'font_name': "DengXian",
	'bold': 1,
	'align': 'center',
	'valign': 'vcenter',
	'bottom': 1,
	'bottom_color': '#555555'
	})
title_with_right_format = workbook.add_format({
	'font_name': "DengXian",
	'bold': 1,
	'align': 'center',
	'valign': 'vcenter',
	'right': 1,
	'right_color': '#555555'
	})
content_format = workbook.add_format({
	'font_name': "DengXian",
	'valign': 'vcenter'
	})
keyword_format = workbook.add_format({
	'font_name': "DengXian",
	'valign': 'vcenter',
	'bold': 1,
	'font_color': 'red'
	})
to_fill_format = workbook.add_format({
	'bg_color': '#DCEDFF',
	'border': 1,
	'border_color': '#CCCCCC'
	})
# border_bottom = workbook.add_format({
# 	'bottom': 1,
# 	'bottom_color': '#AAAAAA'
# 	})

worksheet.merge_range("A1:A2", "段落标号", title_with_bottom_format)
worksheet.merge_range('B1:E1', "涉及行业", title_format)
worksheet.merge_range("F1:F2", "出现次数", title_with_bottom_format)
worksheet.merge_range("G1:G2", "其他关键词（以'、'分隔）", title_with_bottom_format)
worksheet.merge_range("H1:L1", "具体举措", title_format)
worksheet.merge_range("M1:M2", "态度", title_with_bottom_format)
worksheet.merge_range("N1:N2", "位置", title_with_bottom_format)
worksheet.merge_range("O1:O2", "上下文", title_with_bottom_format)
worksheet.write("B2", "2位数代码", title_with_bottom_format)
worksheet.write("C2", "3位数代码", title_with_bottom_format)
worksheet.write("D2", "4位数代码", title_with_bottom_format)
worksheet.write("E2", "行业名称", title_with_bottom_format)
worksheet.write("H2", "产品", title_with_bottom_format)
worksheet.write("I2", "技术", title_with_bottom_format)
worksheet.write("J2", "项目", title_with_bottom_format)
worksheet.write("K2", "园区", title_with_bottom_format)
worksheet.write("L2", "其他", title_with_bottom_format)
curExcelRow = 2

for paragraph in kywdStat:
	if not kywdStat[paragraph]["found_keyword"]:
		worksheet.write(curExcelRow, 0, paragraph, title_with_right_format)
		curExcelRow += 1
	else:
		startExcelRow = curExcelRow
		curIndustryIndexs = kywdStat[paragraph]["related_industry"]
		curFoundKeywords  = kywdStat[paragraph]["statistic_result"]
		for industryIndex in curIndustryIndexs:
			fstIndustry = kywdDict[industryIndex]
			worksheet.write(curExcelRow, 1, fstIndustry["index"], content_format)
			if fstIndustry["name"] in curFoundKeywords:
				worksheet.write(curExcelRow, 4, fstIndustry["name"], keyword_format)
				worksheet.write(curExcelRow, 5, curFoundKeywords[fstIndustry["name"]], content_format)
			else:
				worksheet.write(curExcelRow, 4, fstIndustry["name"], content_format)
			curExcelRow += 1
			secIndustryIndexs = fstIndustry["child"]
			for secIndustryIndex in secIndustryIndexs:
				secIndustry = secIndustryIndexs[secIndustryIndex]
				worksheet.write(curExcelRow, 2, secIndustry["index"], content_format)
				if secIndustry["name"] in curFoundKeywords:
					worksheet.write(curExcelRow, 4, secIndustry["name"], keyword_format)
					worksheet.write(curExcelRow, 5, curFoundKeywords[secIndustry["name"]], content_format)
				else:
					worksheet.write(curExcelRow, 4, secIndustry["name"], content_format)
				trdIndustryIndexs = secIndustry["child"]
				if len(trdIndustryIndexs) < 2:
					worksheet.write(curExcelRow, 3, secIndustry["index"] + "0", content_format)
					curExcelRow += 1
				else:
					curExcelRow += 1
					for trdIndustryIndex in trdIndustryIndexs:
						trdIndustry = trdIndustryIndexs[trdIndustryIndex]
						worksheet.write(curExcelRow, 3, trdIndustry["index"], content_format)
						if trdIndustry["name"] in curFoundKeywords:
							worksheet.write(curExcelRow, 4, trdIndustry["name"], keyword_format)
							worksheet.write(curExcelRow, 5, curFoundKeywords[trdIndustry["name"]], content_format)
						else:
							worksheet.write(curExcelRow, 4, trdIndustry["name"], content_format)
						curExcelRow += 1
		worksheet.merge_range("A" + str(startExcelRow + 1) + ":A" + str(curExcelRow), paragraph, title_with_right_format)

for row in range(2, curExcelRow):
	for col in range(6, 15):
		worksheet.write(row, col, "", to_fill_format)

workbook.close()

