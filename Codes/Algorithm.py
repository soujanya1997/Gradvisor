from collections import defaultdict
import xlrd
import glob
import os
import xlwt
from xlwt import Workbook
import openpyxl
wb1 = xlrd.open_workbook("Jobs.xls")
sheet1 = wb1.sheet_by_index(1)
wb2 = xlrd.open_workbook("Technical_Consolidation.xls")
sheet2 = wb2.sheet_by_index(0)
for i in range(1,sheet1.nrows -1):
	try:
		job[sheet1.cell_value(i,0)] = sheet1.cell_value(i,1)
	except:
		print("",end = "")
	for i in range(2,sheet2.nrows -1):
		try:
			grad[sheet1.cell_value(i,0)] = 1
		except:
			print("",end = "")
for i in range():
	for j in range():
		similarity_matrix[grad[i]][job[j]] = Result[][]
		g = {}
		for x in job:
			g[x] = sorted(similarity_matrix[x].keys(), key=lambda g: similarity_matrix[x][g], reverse = True)
		for x in grad:
			g[x] = sorted(similarity_matrix.keys(), key=lambda g: similarity_matrix[g][x], reverse = True)
		while g:
			d = {}
			for x in grad:
				d[x] = (similarity_matrix[g[x][0]][x] - similarity_matrix[g[x][1]][x])
				s = {}
			for x in job:
				s[x] = (similarity_matrix[x][g[x][0]] - similarity_matrix[x][g[x][1]])
				f = max(d, key=lambda n: d[n])
				t = max(s, key=lambda n: s[n])
				t, f = (f, g[f][0]) if d[f] > s[t] else (g[t][0], t)
				v = min(job[f], grad[t])
				res[f][t] += v
				grad[t] -= v
				if grad[t] == 0:
					for k, n in job.items():
						if n != 0:
							g[k].remove(t)
						del g[t]
					del grad[t]
						job[f] -= v
				if job[f] == 0:
					for k, n in grad.items():
						if n != 0:
							g[k].remove(f)
							del g[f]
					del job[f]
for n in cols:
	print ("\t" , n, end = '')
	print ()
satisfaction = 0
for g in sorted(similarity_matrix):
	print (g, "\t", end = '')
for n in cols:
	y = res[g][n]
if y != 0:
	print (y)
satisfaction += y * similarity_matrix[g][n]
print("\t", end = '')
print()
print ("\n\nTotal satisfaction = ", satisfaction)