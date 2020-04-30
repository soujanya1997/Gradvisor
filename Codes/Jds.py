import xlrd
import glob
import os
import xlwt
from fuzzywuzzy import fuzz
from fuzzywuzzy import process
from xlwt import Workbook
import openpyxl
threshold = 85
Technologies = []
Languages = []
Tools = []
outputFile1 = "Jobs.xls"
wb1 = Workbook()
Jobs = wb1.add_sheet('Jobs')
wb = xlrd.open_workbook("Job_Descriptions.xlsx")
sheet1 = wb.sheet_by_index(0)
sheet2 = wb.sheet_by_index(1)
for j in range(3):
for i in range(sheet1.nrows -1):
if j ==0:
try:
Technologies.append(sheet1.cell_value(i,j))
except:
print("",end = "")
if j ==1:
try:
Tools.append(sheet1.cell_value(i,j))
except:
print("",end = "")
if j ==2:
try:
Languages.append(sheet1.cell_value(i,j))
except:
print("",end = "")
print(Technologies)
print(Tools)
print(Languages)
for i in range(1,sheet2.nrows-1):
JId = sheet2.cell_value(i, 0)
Jobs.write(i,0,JId)
print(JId)
row1 = sheet2.cell_value(i, 18) + " " + sheet2.cell_value(i, 19) + " " + sheet2.cell_value(i,20) + " " + sheet2.cell_value(i,17)
str1 = ""
str2 = ""
str3 = ""
print('TECHS')
for tech in Technologies:
if fuzz.partial_ratio(tech, row1) > threshold:
print(tech + ", ", end = "")
str1 = str1 + tech + ", "
Jobs.write(i, 2, str1)
print()
print('LANGS')
for lang in Languages:
if fuzz.partial_ratio(lang, row1) > threshold:
print(lang + ", ", end ="")
str2 = str2 + lang + ", "
Jobs.write(i, 3, str2)
print()
print('TOOLS')
for tool in Tools:
if fuzz.partial_ratio(tool, row1) > threshold:
print(tool + ", ", end ="")
str3 = str3 + tool + ", "
Jobs.write(i, 4, str3)
print()
print(" - - - - - - - - - - - - - - - - ")
#
wb1.save(outputFile1)