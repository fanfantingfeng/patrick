#coding=GBK

import pandas as pd
from pandas import DataFrame,Series
import numpy as np
import datetime
import os
import xlrd

desk = os.path.expanduser('~') + '\\Desktop'

path = desk + '\\����'

data = DataFrame()

filenames = os.listdir(path)
for i in filenames:
	file_path = path + '\\' + i
	data = data.append(pd.read_excel(file_path))

data.to_excel(desk + "\\�ϲ�����.xlsx", index=False)
print("�ϲ������!")
	









