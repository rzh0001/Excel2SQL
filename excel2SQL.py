#!/usr/bin/python
# -*- coding: utf-8 -*-

# Script Name   : excel2SQL.py
# Author        : Smily
# Created       : 2017-01-11
# Last Modified : 2017-01-13 
# Version       : 0.5

# Modifications	: 

# Description   : This script will generate insert sql accroding to the config and data in the excel file

import sys, os
import xlrd
from datetime import date,datetime
import unicodedata

reload(sys)
sys.setdefaultencoding('utf8')

#获取脚本文件的当前路径
def cur_file_dir():
	#获取脚本路径
	path = sys.path[0]
	#如果是脚本文件，则返回的是脚本的目录，
	if os.path.isdir(path):
		return path
	#如果是py2exe编译后的文件，则返回的是编译后的文件路径
	elif os.path.isfile(path):
		return os.path.dirname(path)
			
if sys.argv[1] == "help" or len(sys.argv) != 3:
	print("Usage:\n\t%s excel_file sheet_index") % (sys.argv[0])
	sys.exit(1)

if os.path.exists(sys.argv[1]):
	filepath = sys.argv[1]
else:
	print("ERROR: File Not Exist [%s]") % ( sys.argv[1] )
	filepath = cur_file_dir() + os.path.sep + sys.argv[1]
	print("Try to find %s") % ( filepath )
	if not os.path.exists(filepath):
		print("ERROR: File Not Exist [%s]") % ( filepath )
		sys.exit(1)
		
sqlfile = filepath[:filepath.rfind('.')] + '.sql'

sheetNumber = int(sys.argv[2])

# 表格第一行配置表名
tableNameRowNumber = 0
# 第二行配置字段名
columnNameRowNumber = 1
# 第三行配置字段类型
colunmnTypeRowNumber = 2
# 数据从第四行开始
firstDataRowNumber = 3

try:
	workbook = xlrd.open_workbook(filepath)
except:
	print("ERROR: Cannot Open [%s]\n Is it a Excel file?") % ( filepath )
	sys.exit(1)
	
try:
	sheet = workbook.sheets()[sheetNumber]
except:
	print("ERROR: Worksheet index [%d] Not Exist") % ( sheetNumber )
	sys.exit(1)

if sheet.cell(tableNameRowNumber,1).ctype != 0:
	tableName = sheet.cell(tableNameRowNumber,1).value
else:
	print("ERROR: TABLE_NAME is null") 
	sys.exit(1)

ncols = sheet.ncols
nrows = sheet.nrows

sqlMode = 0

print '%d LINES INSERT SQL WILL BE GENERATED...' % ( nrows - firstDataRowNumber )
print 'TARGET TABLE_NAME: %s' % (tableName)
print 'TARGET FILE: %s' % (sqlfile)

columnName = sheet.row_values(columnNameRowNumber) 
columnDefind = ''
for i in range(len(columnName)):
	columnDefind = columnDefind + columnName[i] + (', ' if i + 1 < len(columnName) else ' ')

colunmnType = sheet.row_values(colunmnTypeRowNumber)
for i in range(len(colunmnType)):
	if colunmnType[i] == 'function':
		sqlMode = 1
		continue

prefix = 'INSERT ' + tableName + ' (' + columnDefind + ') '
if sqlMode == 0:
	prefix += 'VALUES ('
	suffix = ');\n'
else:
	prefix += 'SELECT '
	suffix = 'FROM DUAL;\n'
	
sql = ''
for i in range(3, nrows):
	tmpStr = ''
	for j in range(ncols):
		cellMode = 0
		if (sheet.cell(i,j).ctype == 0):
			cellValue = 'null'
			cellMode = 1
		else:
			cellValue = sheet.cell_value(i, j)
			if (sheet.cell(i,j).ctype == 3):
				date_value = xlrd.xldate_as_tuple(sheet.cell_value(i,j), workbook.datemode)
				cellValue = date(*date_value[:3]).strftime('%Y%m%d')
			
		if colunmnType[j] == 'int':
			cellMode = 1
			cellValue = str(int(cellValue))
		elif colunmnType[j] == 'float':
			cellMode = 1
			cellValue = str(cellValue)
		elif colunmnType[j] == 'function':
			cellMode = 1
			
		if cellMode == 0:
			tmpStr = tmpStr + "'" + cellValue.replace("'", "''") + "'" + (', ' if j + 1 < ncols else ' ')
		else:
			tmpStr = tmpStr  + cellValue + (', ' if j + 1 < ncols else ' ')
	sql = sql + prefix + tmpStr + suffix
	
with open(sqlfile, 'w') as f:
	    f.write(sql.encode('utf-8'))
f.close()

# XL_CELL_EMPTY	0	empty string u''
# XL_CELL_TEXT	1	a Unicode string
# XL_CELL_NUMBER	2	float
# XL_CELL_DATE	3	float
# XL_CELL_BOOLEAN	4	int; 1 means TRUE, 0 means FALSE
# XL_CELL_ERROR	5	int representing internal Excel codes; for a text representation, refer to the supplied dictionary error_text_from_code
# XL_CELL_BLANK	6	empty string u''. Note: this type will appear only when open_workbook(..., formatting_info=True) is used.


