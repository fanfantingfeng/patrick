#coding=GBK

import pandas as pd
from pandas import DataFrame,Series
import numpy as np
import datetime
import os
import xlrd

desk = os.path.expanduser('~') + '\\Desktop'

path = desk + '\\������ϸ'
filenames = os.listdir(path)

xz = DataFrame()
sl = DataFrame()

for i in filenames:
	filename = path + '\\' + i
	wb = xlrd.open_workbook(filename)
	names = wb.sheet_names()
	for j in names:
		if '������ϸ' in j:
			xz = xz.append(pd.read_excel(filename, sheet_name='������ϸ',header=1))
		if 'Сʱ��' in j:
			sl = xz.append(pd.read_excel(filename, sheet_name='Сʱ��',header=0))

xz = DataFrame(xz[xz['SAP���'].notnull()],columns=['SAP���'])
sl = DataFrame(sl[sl['SAP���'].notnull()],columns=['SAP���'])
comb = xz.append(sl)
comb = comb.drop_duplicates()

comb.loc[:,'SAP���'] = comb.loc[:,'SAP���'].astype('str')
comb = DataFrame(comb[comb['SAP���'].str.isdecimal()])
comb.loc[:,'����'] = comb.loc[:,'SAP���'].apply(lambda x:len(x))
comb = DataFrame(comb[(comb['����'] == 7)&(comb['SAP���'].str.startswith('6'))],columns=['SAP���'])

comb.to_excel(desk + '\\��н����.xlsx', index=False)
os.startfile(desk + '\\��н����.xlsx')
