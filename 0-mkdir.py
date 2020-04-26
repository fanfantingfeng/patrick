#coding=GBK

import pandas as pd
from pandas import DataFrame,Series
import numpy as np
import datetime
import os
import xlrd

print("欢迎使用路径创建模板\n一切解释权均归开发者所有!\n开发者: 凡凡\n")

#参数指定
desk = os.path.join(os.path.expanduser("~"),"Desktop")

path = 'D:\根目录\项目\中间表格\日期参数'
path_mid = 'D:\根目录\项目\中间表格'
path_text = desk + "\\1-text"
path_char = desk + "\\2-split"
path_workbook = desk + "\\工资明细"


def crt_dir(path):
	if os.path.exists(path):
		print("文件路径已存在,无需重新创建!")
	else:
		os.path.makedir(path)
		print("文件路径已创建!")

list_dir = [path, path_mid, path_text, path_char,path_workbook]
for i in list_dir:
	crt_dir(i)

input()
