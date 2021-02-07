#--coding:utf-8--

import pandas as pd
from pandas import DataFrame
import numpy as np
import datetime
import os
import xlrd


#参数指定
desk = os.path.join(os.path.expanduser("~"),"Desktop")

path = 'D:\\根目录\\项目\\中间表格\\日期参数\\'
path_mid = 'D:\\根目录\\项目\\中间表格\\'
path_text = desk + "\\1-text\\"
path_char = desk + "\\2-split\\"

month = (datetime.datetime.now()-datetime.timedelta(datetime.datetime.now().day+1)).month
path_date = path + str(month) + '月开始日期整理表.xlsx'


#检查
if os.path.exists(path_mid + '附加专项.xlsx'):
	df_fj = DataFrame(pd.read_excel(path_mid + '附加专项.xlsx'),columns=['SAP编号','姓名','子女教育','住房租金','住房贷款','赡养老人','继续教育'])
	df_fj = df_fj.melt(id_vars=['SAP编号','姓名'], var_name="属性", value_name="金额")
	df_fj.loc[:,"异常值报告"] = ""
	df_fj.loc[df_fj['金额']>2000, "异常值报告"] = "金额过大,税务系统数据和薪资数据集-SAP数据未对齐,请检查"
	df_fj = DataFrame(df_fj[df_fj['异常值报告']!=""])
	if len(df_fj)>=1:
		print(df_fj)
	else:
		print("未发现错误!")
else:
	print("未发现数据!")

input()
