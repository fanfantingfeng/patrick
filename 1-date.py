#coding=GBK

import pandas as pd
from pandas import DataFrame,Series
import numpy as np
import datetime
import os
import xlrd

print("��ӭʹ�úϲ����ģ��\nһ�н���Ȩ���鿪��������!\n������: ����\n")

#����ָ��
path = 'C:\\Users\\hq01ug601\\Desktop\\������ϸ'
path_mid = 'D:\\��Ŀ¼\\��Ŀ\\�м���'


#���·�Χ����
dir_hrlim = {"���·�Χ":"��Χ","�ܲ���ֱ�ܼ༶����":"SH","�ܲ�MB�ܼ༶����":"SH","�ܲ�MC�ܼ༶����":"SH",
			"�ܲ���ܼ༶����":"SH","�ܲ���ֱ������":"SH","�ܲ�MB������":"SH","�ܲ�MC������":"SH",
			"�ܲ��������":"SH","�ܲ���ֱ���ܼ���":"SH","�ܲ�MB���ܼ���":"SH","�ܲ�MC���ܼ���":"SH",
			"�ܲ�����ܼ���":"SH","�ܲ���ֱԱ��":"SH","�ܲ�MBԱ��":"SH","�ܲ�MCԱ��":"SH",
			"�ܲ��Ա��":"SH","�ܲ���ֱפ��":"SH","�ܲ�MBפ��":"SH","�ܲ�MCפ��":"SH",
			"�ܲ��פ��":"SH","�ܲ�������Ա":"SH","�ܲ�MB����פ��":"SH","�ܲ�MC����פ��":"SH",
			"�ܲ������פ��":"SH","�ܲ���ֱ����פ��":"SH","�Ϻ���ֱ":"SH","�Ϻ�MB":"SH",
			"�Ϻ�MC":"SH","������ֱ":"SU","����MB":"SU","����MC":"SU","�Ͼ���ֱ":"NJ",
			"�Ͼ�MB":"NJ","�Ͼ�MC":"NJ","�Ϸ���ֱ":"HF","�Ϸ�MB":"HF","�Ϸ�MC":"HF",
			"������ֱ":"HZ","����MB":"HZ","����MC":"HZ","������ֱ":"NB","����MB":"NB",
			"����MC":"NB","������ֱ":"WZ","����MB":"WZ","����MC":"WZ","������������":"WZ",
			"������ֱ":"BJ","����MB":"BJ","����MC":"BJ","�����ֱ":"TJ","���MB":"TJ",
			"���MC":"TJ","�����������":"TJ","������ֱ":"JN","����MB":"JN","����MC":"JN",
			"��������ֱ":"HE","������MB":"HE","������MC":"HE","������ֱ":"CC","����MB":"CC",
			"����MC":"CC","������ֱ":"SY","����MB":"SY","����MC":"SY","������������":"SY",
			"̫ԭ��ֱ":"TY","̫ԭMB":"TY","̫ԭMC":"TY","ʯ��ׯ��ֱ":"SJ","ʯ��ׯMB":"SJ",
			"ʯ��ׯMC":"SJ","֣����ֱ":"ZZ","֣��MB":"ZZ","֣��MC":"ZZ","������ֱ":"SX",
			"����MB":"SX","����MC":"SX","������������":"SX","������ֱ":"LZ","����MB":"LZ",
			"����MC":"LZ","��³ľ����ֱ":"WQ","��³ľ��MB":"WQ","��³ľ��MC":"WQ","�ɶ���ֱ":"CD",
			"�ɶ�MB":"CD","�ɶ�MC":"CD","�ɶ���������":"CD","������ֱ":"CQ","����MB":"CQ",
			"����MC":"CQ","������ֱ":"KM","����MB":"KM","����MC":"KM","������������":"GZ",
			"������ֱ":"GZ","����MB":"GZ","����MC":"GZ","������ֱ":"SZ","����MB":"SZ",
			"����MC":"SZ","������ֱ":"NN","����MB":"NN","����MC":"NN","�人��ֱ":"WH",
			"�人MB":"WH","�人MC":"WH","�人��������":"WH","�ϲ���ֱ":"NC","�ϲ�MB":"NC",
			"�ϲ�MC":"NC","������ֱ":"FZ","����MB":"FZ","����MC":"FZ","��ݸ��ֱ":"DG",
			"��ݸMB":"DG","��ݸMC":"DG","��ɳ��ֱ":"CS","��ɳMB":"CS","��ɳMC":"CS",
			"������ֱ":"GY","����MB":"GY","����MC":"GY","�ൺ��ֱ":"QD","�ൺMB":"QD",
			"�ൺMC":"QD","���ɹ���ֱ":"NM","���ɹ�MB":"NM"}

#�����Բ���


root = 'D:\\��Ŀ¼\\�����춯\\'
month = (datetime.datetime.now() - datetime.timedelta(30,0,0,0)).month
file800 = root + str(month) + "�������춯��-800.xlsx"
file830 = root + str(month) + "�������춯��-830.xlsx"

if (os.path.exists(file800)) & (os.path.exists(file800)):
	df = pd.read_excel(file800).append(pd.read_excel(file830))
elif os.path.exists(file800):
	df = pd.read_excel(file800)
elif os.path.exists(file830):
	df = pd.read_excel(file830)
else:
	df = DataFrame()
	print("δ�ҵ�������������춯��!")
	
df = df.reset_index()

def stan(data):
	for i in range(len(data.index)):
		data.loc[i,'�³�����'] = datetime.datetime(int(data.loc[i,'��']),int(data.loc[i,'��']),1)
		data.loc[i,'��ĩ����'] = datetime.datetime(int(data.loc[i,'��']),int(data.loc[i,'��'])+1,1) - datetime.timedelta(1,0,0,0)
		data.loc[i,'��ʼ����'] = data.loc[i,'��ĩ����'].strftime("%Y%m%d")
		if data.loc[i,'��ְ'] == '��Ա����ְ':
			data.loc[i,'��ʼ����'] = data.loc[i,'��ְ����'].strftime("%Y%m%d")
		else:
			data.loc[i,'��ʼ����'] = data.loc[i,'�³�����'].strftime("%Y%m%d")
	data.loc[:,'��������'] = "9991231"
	return data

if len(df) >= 1:
	hrlim = DataFrame(Series(dir_hrlim),columns=['��Χ'])
	stan(df)
	data = pd.merge(df, hrlim, left_on='���·�Χ����', right_index=True, how='left')
	data.loc[:,'ϵͳ'] = 800
	data.loc[data['SAP��Ա���'] > 6000000,"ϵͳ"] = 830
	data_end = data.loc[:,["SAP��Ա���","��ʼ����","��ʼ����","��������","��Χ","ϵͳ"]]
	file_name = 'D:\\��Ŀ¼\\��Ŀ\\�м���\\���ڲ���\\' + str(month) +'�¿�ʼ���������.xlsx'
	data_end.to_excel(file_name, index=False)
	print("��ʼ����������ѳɹ�����!")
else:
	print("��ʼ���������δ����,��˶������춯����Ϣ!")

input()













