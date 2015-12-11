#!/usr/bin
#coding:utf-8
import sys, codecs
from mmap import mmap, ACCESS_READ
from xlrd import open_workbook

def exportFile(fileName, data):
	f = codecs.open(fileName,'a','utf-8')
	f.write(unicode(data)+'\n')
	f.close

if __name__ == '__main__':
	inputFileName = sys.argv[1]
	#############################
	wb = open_workbook(inputFileName)

	sheets = wb.sheets()
	for i in xrange(5,len(sheets)):
		print 'Sheet: ',sheets[i].name
		for row in xrange(sheets[i].nrows):
			values = []
			for col in xrange(sheets[i].ncols):
				data = sheets[i].cell(row,col).value
				values.append(data)
			exportFile(sheets[i].name+'.xlsx',values)
		break