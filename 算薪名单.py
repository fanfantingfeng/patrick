#coding=GBK

import pandas as pd
from pandas import DataFrame,Series
import numpy as np
import datetime
import os
import xlrd

desk = os.path.expanduser('~') + '\\Desktop'

path = desk + '\\工资明细'
filenames = os.listdir(path)

xz = DataFrame()
sl = DataFrame()

for i in filenames:
	filename = path + '\\' + i
	wb = xlrd.open_workbook(filename)
	names = wb.sheet_names()
	for j in names:
		if '工资明细' in j:
			xz = xz.append(pd.read_excel(filename, sheet_name='工资明细',header=1))
		if '小时工' in j:
			sl = xz.append(pd.read_excel(filename, sheet_name='小时工',header=0))

xz = DataFrame(xz[xz['SAP编号'].notnull()],columns=['SAP编号'])
sl = DataFrame(sl[sl['SAP编号'].notnull()],columns=['SAP编号'])
comb = xz.append(sl)
comb = comb.drop_duplicates()

comb.loc[:,'SAP编号'] = comb.loc[:,'SAP编号'].astype('str')
comb = DataFrame(comb[comb['SAP编号'].str.isdecimal()])
comb.loc[:,'长度'] = comb.loc[:,'SAP编号'].apply(lambda x:len(x))
comb = DataFrame(comb[(comb['长度'] == 7)&(comb['SAP编号'].str.startswith('6'))],columns=['SAP编号'])

comb.to_excel(desk + '\\算薪名单.xlsx', index=False)
os.startfile(desk + '\\算薪名单.xlsx')
