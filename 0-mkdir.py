#coding=GBK

import pandas as pd
from pandas import DataFrame,Series
import numpy as np
import datetime
import os
import xlrd

print("��ӭʹ��·������ģ��\nһ�н���Ȩ���鿪��������!\n������: ����\n")

#����ָ��
desk = os.path.join(os.path.expanduser("~"),"Desktop")

path = 'D:\��Ŀ¼\��Ŀ\�м���\���ڲ���'
path_mid = 'D:\��Ŀ¼\��Ŀ\�м���'
path_text = desk + "\\1-text"
path_char = desk + "\\2-split"
path_workbook = desk + "\\������ϸ"


def crt_dir(path):
	if os.path.exists(path):
		print("�ļ�·���Ѵ���,�������´���!")
	else:
		os.path.makedir(path)
		print("�ļ�·���Ѵ���!")

list_dir = [path, path_mid, path_text, path_char,path_workbook]
for i in list_dir:
	crt_dir(i)

input()
