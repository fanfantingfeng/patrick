#coding=GBK

import pandas as pd
from pandas import DataFrame
import numpy as np
import os
import xlrd

print("��ӭʹ�úϲ����ģ��\nһ�н���Ȩ���鿪��������!\n������: ����\n")

#����ָ��
desk = os.path.join(os.path.expanduser("~"),"Desktop")
path = desk + "\\������ϸ"
path_mid = 'D:\\��Ŀ¼\\��Ŀ\\�м���'
print(path)
sheet_dir = {"����":"kq","����":"jt","�籣":"ss",
			 "����-˰��":"sw","����-SAP":"sap",
			 "н���춯��":"xz","�����춯��":"jt1",
			 "Сʱ��":"sl","������ϸ":"bk"}

col_kq = ["SAP���","����","���¼�н����","Ӧ��������","ʵ�ʳ�������",
			"ȫ�ڽ�","����","������","��������","��н�¼�����","��������",
			"�������","ɥ������","��������","��������","��������","���������",
			"���˼�����","����Сʱ","��������","��н�¼�����","����δ�򿨣��Σ�",
			"�ǵ���δ�򿨣��Σ�","�ǵ��̳ٵ�0-30M","�ǵ��̳ٵ�31-60M",
			"�ǵ��̳ٵ�61-120M","�ǵ��̳ٵ�120M����","��������30M����",
			"��������30M����","���̳ٵ�0-10M","���̳ٵ�11-30M",
			"���̳ٵ�31-60M","���̳ٵ�61-120M","���̳ٵ�120M����",
			"��������1Сʱ��","��������1Сʱ����",
			"ƽʱ�Ӱ�ʱ","���ռӰ�ʱ","��ĩ�Ӱ�ʱ"]
col_jt = ["SAP���","����","������","���"]

col_ss = ["SAP���","����","�籣�˻�","�������˻�",
		"���Ͻ��ɽ�����","���Ͻ��ɽ�˾","ҽ�ƽ��ɽ�����","ҽ�ƽ��ɽ�˾",
		"ʧҵ���ɽ�����","ʧҵ���ɽ�˾","���˽��ɽ�˾","�������ɽ�˾",
		"��������ɽ�����","��������ɽ�˾","���ϲ��ɽ�����","���ϲ��ɽ�˾",
		"ҽ�Ʋ��ɽ�����","ҽ�Ʋ��ɽ�˾","ʧҵ���ɽ�����","ʧҵ���ɽ�˾",
		"���˲��ɽ�˾","�������ɽ�˾","�����𲹽ɽ�����","�����𲹽ɽ�˾",
		"����ҽ�ƽ��ɽ�����","����ҽ�ƽ��ɽ�˾",
		"����ҽ�Ʋ��ɽ�����","����ҽ�Ʋ��ɽ�˾"]


#���ݳ�ʼ��
mid_file = os.listdir(path_mid)
if len(mid_file) >= 1:
	for i in mid_file:
		filename = path_mid + "\\" + i
		if os.path.isfile(filename):
			os.remove(filename)
	print("��ʼ�������!")
else:
	print("���������ʼ��!")



#�Զ��庯��
def output(data,text):
	data.dropna(axis=0,how='all',inplace=True)
	data.dropna(axis=1,how='all',inplace=True)
	data = data.drop_duplicates()
	data = data.replace(" ", 0)
	data.fillna(0,inplace=True)
	df = DataFrame(data[data.loc[:,'SAP���'].notnull()])
	try:
		df.loc[:,'SAP���'] = df.loc[:,'SAP���'].astype('int')
	except ValueError:
		df.loc[:,'SAP���'] = df.loc[:,'SAP���'].astype('str')
	if len(df.index) >= 1:
		df.loc[:,'SAP���'] = df.loc[:,'SAP���'].astype('str')
		df = df[(df.loc[:,'SAP���'].notnull())&(df.loc[:,'SAP���'].str.isnumeric())]
		df = df.fillna(0)
		df.loc[:,'SAP���'] = df.loc[:,'SAP���'].astype('int')
		filename = path_mid + "\\" + text + '.xlsx'
		if len(df.index) >= 1:
			df.to_excel(filename, index=False)
			print(text + "��������,����Ϊ: " + str(len(df)))
		else:
			print(text + "�������赼��!")
	else:
		print(text + "����δ����!")

def stan(data):
	data.dropna(axis=0,how='all',inplace=True)
	data.dropna(axis=1,how='all',inplace=True)
	data.loc[:,'SAP���'] = data.loc[:,'SAP���'].astype('str')
	data = data[(data.loc[:,'SAP���'].notnull())&(data.loc[:,'SAP���'].str.isnumeric())]
	data = data.fillna(0)
	data.loc[:,'SAP���'] = data.loc[:,'SAP���'].astype('int')
	return data


#��������
filename = os.listdir(path)

kq = DataFrame()
jt = DataFrame()
ss = DataFrame()
sw = DataFrame()
sap = DataFrame()

xz = DataFrame()
jt1 = DataFrame()
sl = DataFrame()
bk = DataFrame()
		
		
for i in filename:
	files = path + "\\" + i
	wb = xlrd.open_workbook(files)
	names = wb.sheet_names()
	for j in names:
		if "����ͳ��" in j:
			kq = kq.append(pd.read_excel(files,sheet_name="����ͳ��",header=1),ignore_index=True)
		if "������ϸ" in j:
			jt = jt.append(pd.read_excel(files,sheet_name="������ϸ"),ignore_index=True)
		if "�籣ͳ��" in j:
			ss = ss.append(pd.read_excel(files,sheet_name="�籣ͳ��",header=3),ignore_index=True)
		if "˰��ϵͳ" in j:
			sw = sw.append(pd.read_excel(files,sheet_name="ר��ӿ۳�-˰��ϵͳ"),ignore_index=True)
		if "н�����ݼ�" in j:
			sap = sap.append(pd.read_excel(files,sheet_name="н�����ݼ�-Sap"),ignore_index=True)
		if "н���춯��" in j:
			xz = xz.append(pd.read_excel(files,sheet_name="н���춯��"),ignore_index=True)
		if "�����춯��" in j:
			jt1 = jt1.append(pd.read_excel(files,sheet_name="�����춯��"),ignore_index=True)
		if "Сʱ��" in j:
			sl = sl.append(pd.read_excel(files,sheet_name="Сʱ��"),ignore_index=True)
		if "����" in j:
			bk = bk.append(pd.read_excel(files,sheet_name="������ϸ",dtype={'���д���':'str','�����˺�':'str'}),ignore_index=True)





#����
kq = DataFrame(kq, columns=col_kq)
output(kq,"����")

#������ϸ
jt = DataFrame(jt, columns=col_jt)
output(jt,"������ϸ")

#�籣
ss = DataFrame(ss, columns=col_ss)
output(ss,"�籣")

#����ר��
stan(sw)
stan(sap)
if (len(sw.index) >= 1)&(len(sap.index) >= 1):
	fj = pd.merge(sw,sap,on='SAP���',how='outer')
	fj.loc[:,'��Ů����'] = fj.loc[:,'�ۼ���Ů����_x'] - fj.loc[:,'�ۼ���Ů����_y']
	fj.loc[:,'ס�����'] = fj.loc[:,'�ۼ�ס�����_x'] - fj.loc[:,'�ۼ�ס�����_y']
	fj.loc[:,'ס������'] = fj.loc[:,'�ۼ�ס������_x'] - fj.loc[:,'�ۼ�ס������_y']
	fj.loc[:,'��������'] = fj.loc[:,'�ۼ���������_x'] - fj.loc[:,'�ۼ���������_y']
	fj.loc[:,'��������'] = fj.loc[:,'�ۼƼ�������_x'] - fj.loc[:,'�ۼƼ�������_y']
	fj = DataFrame(fj,columns=['SAP���','��Ů����','ס�����','ס������','��������','��������'])
	output(fj,"����ר��")
else:
	print("δ���������ĸ���ר���������!")

#н���춯��
if len(xz) >= 1:
	xz = DataFrame(xz,columns=['SAP���','����','������͹��ʱ�׼','н��'])
	output(xz,"н���춯��")
else:
	print("δ����н���춯����!")

#�����춯��
if len(jt1) >= 1:
	jt1 = DataFrame(jt1,columns=['SAP���','����','��Ŀ','���'])
	output(jt1,"�����춯��")
else:
	print("δ���ֽ����춯����!")

#Сʱ��
if len(sl) >= 1:
	sl = DataFrame(sl,columns=['SAP���','Сʱ��','ʱн','����','��н','���','ʧ��','����˰'])
	output(stan(sl),"Сʱ��")
else:
	print("δ����Сʱ������!")

#������ϸ
if len(bk) >= 1:
	bk = DataFrame(bk,columns=['SAP���','���д���','�����˺�'])
	bk = bk.dropna(axis=0,how='any')
	bk.loc[:,'���д���'] = bk.loc[:,'���д���'].astype('str')
	bk.loc[:,'�����˺�'] = bk.loc[:,'�����˺�'].astype('str')
	output(bk,"������ϸ")
else:
	print("δ����������ϸ����!")

print("\n�м�������ɴ���,�������н����,лл!")
input()


