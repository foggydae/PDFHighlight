# coding: utf-8

import re
import json
import jieba
import nltk
from pprint import pprint

with open("key_word_dict/dict_with_alias.json") as jsonFile:
	kywdDict = json.load(jsonFile)

stopWords = ['、', '及', '和', '与', '（', '）', '(', ')']
stopWords_split = '、|及|和|与|（|）|(|)'

wordFrequency = []
aliasDict = {}

# Split with stop words. For each segment, cut them up with JIEBA.

def combo_alias(inputList):
	outputList = []
	for i in range(len(inputList) - 1):
		newWord = inputList[i]
		outputList.append(newWord)
		for j in range(i + 1, len(inputList)):
			newWord += inputList[j]
			outputList.append(newWord)
	outputList.append(inputList.pop())
	return outputList

def initialize_alias(inputName):
	outputAlias = {
		'currentCut': {},
		'cutGroup': []
	}
	basicSegs = list(re.split(stopWords_split, inputName))
	for basicSeg in basicSegs:
		if basicSeg is None or basicSeg == '':
			pass
		elif len(basicSeg) <= 2:
			wordFrequency.append(basicSeg)
			if basicSeg not in outputAlias['currentCut']:
				outputAlias['currentCut'][basicSeg] = 1
		else:
			jiebaSegs = list(jieba.cut(basicSeg, cut_all = False))
			outputAlias['cutGroup'].append(jiebaSegs)
			for jiebaSeg in combo_alias(jiebaSegs):
				wordFrequency.append(jiebaSeg)
				if jiebaSeg not in outputAlias['currentCut']:
					outputAlias['currentCut'][jiebaSeg] = 1
	return outputAlias

for secndKeyword in kywdDict:
	aliasDict[secndKeyword] = initialize_alias(kywdDict[secndKeyword]['name'])
	for thirdKeyword in kywdDict[secndKeyword]['child']:
		aliasDict[thirdKeyword] = initialize_alias(kywdDict[secndKeyword]['child'][thirdKeyword]['name'])
		if not len(kywdDict[secndKeyword]['child'][thirdKeyword]['child']) == 1:
			for forthKeyword in kywdDict[secndKeyword]['child'][thirdKeyword]['child']:
				aliasDict[forthKeyword] = initialize_alias(kywdDict[secndKeyword]['child'][thirdKeyword]['child'][forthKeyword]['name'])

fredist = nltk.FreqDist(wordFrequency)
pprint(fredist)

def isUnique (thisIndex, inputWords):
	output = {
		'uniqueWords': [],
		'dupWords': []
	}
	for word in inputWords:
		if fredist[word] == 1:
			output['uniqueWords'].append(word)
			continue
		else:
			count = 0
			for index in aliasDict:
				if not re.match(thisIndex, index):
					continue
				if word in aliasDict[index]['currentCut']:
					count += 1
			if count == fredist[word]:
				output['uniqueWords'].append(word)
			elif fredist[word] - count > 3:
				output['dupWords'].append(word)
	return output

for index in aliasDict:
	checkUnique = isUnique(index, list(aliasDict[index]['currentCut'].keys()))
	for word in checkUnique['uniqueWords']:
		aliasDict[index]['currentCut'][word] = 2
	for word in checkUnique['dupWords']:
		aliasDict[index]['currentCut'][word] = 0

pprint(aliasDict)


for secndKeyword in kywdDict:
	kywdDict[secndKeyword]['alias'] = aliasDict[kywdDict[secndKeyword]['index']]['currentCut']

	for thirdKeyword in kywdDict[secndKeyword]['child']:
		kywdDict[secndKeyword]['child'][thirdKeyword]['alias'] = aliasDict[kywdDict[secndKeyword]['child'][thirdKeyword]['index']]['currentCut']

		if not len(kywdDict[secndKeyword]['child'][thirdKeyword]['child']) == 1:
			for forthKeyword in kywdDict[secndKeyword]['child'][thirdKeyword]['child']:
				kywdDict[secndKeyword]['child'][thirdKeyword]['child'][forthKeyword]['alias'] = aliasDict[kywdDict[secndKeyword]['child'][thirdKeyword]['child'][forthKeyword]['index']]['currentCut']

with open("dict_with_alias_CUT_COMBO.json", "w") as jsonFile:
	jsonFile.write(json.dumps(kywdDict, indent = 2, ensure_ascii = False))

