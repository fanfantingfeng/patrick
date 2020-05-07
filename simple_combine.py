#coding=GBK

import pandas as pd
from pandas import DataFrame,Series
import numpy as np
import datetime
import os
import xlrd

desk = os.path.expanduser('~') + '\\Desktop'

path = desk + '\\汇总'

data = DataFrame()

filenames = os.listdir(path)
for i in filenames:
	file_path = path + '\\' + i
	data = data.append(pd.read_excel(file_path))

data.to_excel(desk + "\\合并数据.xlsx", index=False)
print("合并已完成!")
	









