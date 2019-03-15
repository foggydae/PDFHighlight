# coding: utf-8

import re
import os
import json
import xlsxwriter
import pypandoc
from pprint import pprint

inputFileName = 'dict_with_alias_CUT_COMBO'

with open(inputFileName + ".json") as jsonFile:
	kywdDict = json.load(jsonFile)


workbook = xlsxwriter.Workbook(inputFileName + ".xlsx")
worksheet = workbook.add_worksheet()

worksheet.set_column('A:Z', 10)

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
recommend_format = workbook.add_format({
	'font_name': "DengXian",
	'valign': 'vcenter',
	'bold': 1,
	'font_color': 'red'
	})
unrecommend_format = workbook.add_format({
	'font_name': "DengXian",
	'valign': 'vcenter',
	'font_color': 'grey'
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

worksheet.write("A1", "2位数代码", title_with_bottom_format)
worksheet.write("B1", "3位数代码", title_with_bottom_format)
worksheet.write("C1", "4位数代码", title_with_bottom_format)
worksheet.write("D1", "行业名称", title_with_bottom_format)
worksheet.write("E1", "别名", title_with_bottom_format)

curExcelRow = 1

def write_alias(aliasDict, curRow):

	sortedDictIndex = sorted(aliasDict.items(), key = lambda x: x[1], reverse = True)
	i = 0
	for keys in sortedDictIndex:
		if aliasDict[keys[0]] == 2:
			worksheet.write(curRow, 4 + i, keys[0], recommend_format)
		elif aliasDict[keys[0]] == 1:
			worksheet.write(curRow, 4 + i, keys[0], content_format)
		else:
			worksheet.write(curRow, 4 + i, keys[0], unrecommend_format)
		i += 1



for industryIndex in kywdDict:
	fstIndustry = kywdDict[industryIndex]
	worksheet.write(curExcelRow, 0, fstIndustry["index"], content_format)
	worksheet.write(curExcelRow, 3, fstIndustry["name"], content_format)
	write_alias(fstIndustry['alias'], curExcelRow)

	curExcelRow += 1
	secIndustryIndexs = fstIndustry["child"]
	for secIndustryIndex in secIndustryIndexs:
		secIndustry = secIndustryIndexs[secIndustryIndex]
		worksheet.write(curExcelRow, 1, secIndustry["index"], content_format)
		worksheet.write(curExcelRow, 3, secIndustry["name"], content_format)
		write_alias(secIndustry['alias'], curExcelRow)

		trdIndustryIndexs = secIndustry["child"]
		if len(trdIndustryIndexs) < 2:
			worksheet.write(curExcelRow, 2, secIndustry["index"] + "0", content_format)
			curExcelRow += 1
		else:
			curExcelRow += 1
			for trdIndustryIndex in trdIndustryIndexs:
				trdIndustry = trdIndustryIndexs[trdIndustryIndex]
				worksheet.write(curExcelRow, 2, trdIndustry["index"], content_format)
				worksheet.write(curExcelRow, 3, trdIndustry["name"], content_format)
				write_alias(trdIndustry['alias'], curExcelRow)

				curExcelRow += 1

workbook.close()

