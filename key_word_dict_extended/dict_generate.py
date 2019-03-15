# coding: utf-8

import re
import xlrd
import json
from pprint import pprint

def get_industry_from_index (indexs):
	outputIndustries = []
	for index in indexs:
		if index[0:2] not in outputIndustries:
			outputIndustries.append(index[0:2])
	return outputIndustries

def check_equal_for_array (array1, array2):
	if len(set(array1) & set(array2)) == 0:
		return False
	return True

def clean_id_list (array):
	outputList = []
	if len(array) == 1:
		return array
	array = sorted(array, key=len, reverse=False)
	for i in range(len(array)):
		appendFlag = True
		for j in range(i + 1, len(array)):
			if re.match(array[i], array[j]):
				appendFlag = False
				break
		if appendFlag:
			outputList.append(array[i])
	return outputList

# kywdFile = xlrd.open_workbook("dict_with_alias_0513.xlsx")
kywdFile = xlrd.open_workbook("stop_words_list_0526.xlsx")
kywdTable = kywdFile.sheets()[0]
kywdRows = kywdTable.nrows
kywdCols = kywdTable.ncols

kywdDict = {}
kywdRef = {}
kywdDup = {}

curGrandFather = ""
curFather = ""
curSon = ""

for rowIndex in range(1, kywdRows):
	rowValue = kywdTable.row_values(rowIndex)

	if not rowValue[0] == "":
		curGrandFather = rowValue[0]
		kywdDict[curGrandFather] = {
			"index": rowValue[0],
			"name": rowValue[3],
			"child": {}
		}
		thisIndex = rowValue[0]
	elif not rowValue[1] == "":
		curFather = rowValue[1]
		kywdDict[curGrandFather]["child"][curFather] = {
			"index": rowValue[1],
			"name": rowValue[3],
			"child": {}
		}
		thisIndex = rowValue[1]
		if not rowValue[2] == "":
			curSon = rowValue[2]
			kywdDict[curGrandFather]["child"][curFather]["child"][curSon] = {
				"index": rowValue[2],
				"name": rowValue[3],
			}			
	elif not rowValue[2] == "":
		curSon = rowValue[2]
		kywdDict[curGrandFather]["child"][curFather]["child"][curSon] = {
			"index": rowValue[2],
			"name": rowValue[3],
		}
		thisIndex = rowValue[2]
	else:
		continue

	for colIndex in range(3, kywdCols):
		if rowValue[colIndex] == "":
			continue

		if (not colIndex == 3) and (rowValue[colIndex] == rowValue[3]):
			continue

		if rowValue[colIndex] not in kywdRef:
			kywdRef[rowValue[colIndex]] = [thisIndex]
		else:
			kywdRef[rowValue[colIndex]].append(thisIndex)


# print("\n\n---------------------------------------")
for keyword in kywdRef:
	kywdRef[keyword] = list(set(kywdRef[keyword]))
	kywdRef[keyword] = clean_id_list(kywdRef[keyword])

	# industryNum = kywdRef[keyword][0][0:2]
	# for indexNum in kywdRef[keyword]:
	# 	if not indexNum[0:2] == industryNum:
	# 		print(industryNum)
	# 		print(keyword, "has conflict reference.")

# kywdList = kywdRef.keys()
# kywdList = sorted(kywdList, key=len, reverse=False)

# for thisIndex in range(len(kywdList)):
# 	for thatIndex in range(thisIndex, len(kywdList)):
# 		thisWord = kywdList[thisIndex]
# 		thatWord = kywdList[thatIndex]
# 		if thisWord == thatWord:
# 			continue

# 		if re.search(thisWord, thatWord):
# 			if check_equal_for_array(get_industry_from_index(kywdRef[thisWord]), get_industry_from_index(kywdRef[thatWord])):
# 				# if len(get_industry_from_index(kywdRef[thisWord])) < len(get_industry_from_index(kywdRef[thatWord])):
# 				print("\n")
# 				print(thisWord + ":", ",".join(kywdRef[thisWord]))
# 				print(thatWord + ":", ",".join(kywdRef[thatWord]))
# 			if thisWord not in kywdDup:
# 				kywdDup[thatWord] = [thisWord]
# 			else:
# 				kywdDup[thatWord].append(thisWord)

# print("\n-----------------------------kywdDict-----------------------------")
# pprint(kywdDict)
# with open("structured_dict.json", "w") as jsonFile:
# 	jsonFile.write(json.dumps(kywdDict, indent = 2, ensure_ascii = False))

# print("\n-----------------------------kywdRef -----------------------------")
# pprint(kywdRef)
# with open("reference_dict.json", "w") as jsonFile:
# 	jsonFile.write(json.dumps(kywdRef, indent = 2, ensure_ascii = False))
with open("stop_word_dict.json", "w") as jsonFile:
	jsonFile.write(json.dumps(kywdRef, indent = 2, ensure_ascii = False))

# # print("\n-----------------------------kywdDup -----------------------------")
# # pprint(kywdDup)
# with open("containing_dict.json", "w") as jsonFile:
# 	jsonFile.write(json.dumps(kywdDup, indent = 2, ensure_ascii = False))
