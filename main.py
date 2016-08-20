#!/usr/bin
#coding:utf-8
import sys, codecs, os
from mmap import mmap, ACCESS_READ
from xlrd import open_workbook
from xlwt import Workbook

def convertDataLine(dataArray):
	resultData = ""
	for i in xrange(len(dataArray)):
		resultData += (unicode(dataArray[i]).strip()+', ')
	return resultData[:-2]

def writeLineInFile(fileName, data, numberOfLine):
	f = codecs.open(fileName+'.txt','a','utf-8')
	for i in xrange(numberOfLine):
		f.write(data+'\n')
		#f.write('\n')
	f.close

def writeBook(book, counterDict):
	fileList = []
	for sheetName in book.keys():
		fileList.append(sheetName)
		################################
		counter = 0
		sheetData = book[sheetName]
		index = counterDict[sheetName]
		for i in xrange(len(sheetData)):
			if i == 0:
				writeLineInFile(sheetName, convertDataLine(sheetData[i]), 1)
				#writeLineInFile(sheetName, str(sheetData[i]), 1)
			else:
				stt = int(sheetData[i][index])
				writeLineInFile(sheetName, convertDataLine(sheetData[i]), stt - counter)
				#writeLineInFile(sheetName, str(sheetData[i]), stt - counter)
				counter = 0
	return fileList

def makeCounterDict(book, counterColName):
	counterDict = {}
	for sheetName in book.keys():
		titleList = book[sheetName][0]
		if titleList.index(counterColName) != 0:
			counterDict[sheetName] = titleList.index(counterColName)
	return counterDict

def readInputFile(inputFileName):
	wb = open_workbook(inputFileName)
	########################################
	sheets = wb.sheets()
	book = {}
	for i in xrange(len(sheets)):
		sheetName = sheets[i].name
		sheetValues = []
		for row in xrange(sheets[i].nrows):
			rowValues = []
			for col in xrange(sheets[i].ncols):
				cellData = sheets[i].cell(row,col).value
				rowValues.append(cellData)
			sheetValues.append(rowValues)
		book[sheetName] = sheetValues
	return book

def createBarcode(previousBarcode):
	for i in xrange(len(previousBarcode)):
		try:
			num = int(previousBarcode[i:])
			head = previousBarcode[0:i]
			break
		except:
			continue
	return head+str(num+1)

def writeBarcode(barcodeHead, counterColName):
	f = codecs.open(barcodeHead+'.txt','r','utf-8')
	f1 = codecs.open(barcodeHead+'_addBar.txt','a','utf-8')
	barcode = barcodeHead+'0000'
	for line in f:
		if line.find(counterColName) > 0:
			f1.write(line[:-1]+', Barcode\n')
			continue
		else:
			barcode = createBarcode(barcode)
			f1.write(line[:-1]+', '+barcode+'\n')
	f1.close
	f.close

def exportExcelFile(inputFileName, fileList, counterColName):
	excel = Workbook()
	outputFileName = inputFileName.replace('.xlsx','').replace('.xls','')+'_result.xls'
	for fileName in fileList:
		writeBarcode(fileName, counterColName)
		sheet = excel.add_sheet(fileName)
		writeSheet(sheet, fileName)
	excel.save(outputFileName)
	print "Output now available in "+outputFileName

def writeSheet(sheet, fileName):
	f = codecs.open(fileName+'_addBar.txt','r','utf-8')
	print fileName+'_addBar.txt'
	row = 0
	for line in f:
		row += 1
		col = 0
		dataList = line.split(',')
		for data in dataList:
			sheet.write(row, col, data.strip())
			col += 1
	f.close

def removeFile(fileList):
	for fileName in fileList:
		os.remove(fileName+'.txt')
		os.remove(fileName+'_addBar.txt')

if __name__ == '__main__':
	inputFileName = sys.argv[1]
	counterColName = raw_input("Sort by: ").decode('utf-8')
	#############################
	book = readInputFile(inputFileName)
	counterDict = makeCounterDict(book, counterColName)
	fileList = writeBook(book, counterDict)
	# writeBarcode('KO15061', counterColName)
	exportExcelFile(inputFileName, fileList, counterColName)
	removeFile(fileList)
	#############################
