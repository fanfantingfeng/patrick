#coding=GBK

import pandas as pd
from pandas import DataFrame,Series
import numpy as np
import datetime
import os
import xlrd

print("��ӭʹ�úϲ����ģ��\nһ�н���Ȩ���鿪��������!\n������: ����\n")

#����ָ��
desk = os.path.join(os.path.expanduser("~"),"Desktop")

path = 'D:\\��Ŀ¼\\��Ŀ\\�м���\\���ڲ���\\'
path_mid = 'D:\\��Ŀ¼\\��Ŀ\\�м���\\'
path_text = desk + "\\1-text\\"
path_char = desk + "\\2-split\\"

month = (datetime.datetime.now() - datetime.timedelta(28,0,0,0)).month
path_date = path + str(month) + '�¿�ʼ���������.xlsx'

df_date = pd.read_excel(path_date)


#���������
dir_or = {"����ҽ�ƽ��ɽ�����":"0917","����ҽ�ƽ��ɽ�˾":"0918","���ϲ��ɽ�����":"0919",
		  "���ϲ��ɽ�˾":"0920","ҽ�Ʋ��ɽ�����":"0921","ҽ�Ʋ��ɽ�˾":"0922",
		  "ʧҵ���ɽ�����":"0923","ʧҵ���ɽ�˾":"0924","���˲��ɽ�˾":"0926",
		  "�������ɽ�˾":"0928","�����𲹽ɽ�����":"0929","�����𲹽ɽ�˾":"0930",
		  "����ҽ�Ʋ��ɽ�����":"0931","����ҽ�Ʋ��ɽ�˾":"0932","�ۺϱ��չ�˾���ɽ��":"0937",
		  "�ۺϱ��ո��˲��ɽ��":"0939","Сʱ������":"1004","����":"1202","��������":"1211",
		  "������Ů��":"1213","ֵ�����":"1216","����1":"1217","����2":"1218","ȫ�ڽ�":"1304",
		  "����֮��":"1306","����֮��":"1307","���ڷ���":"1311","��������":"1314","ҵ�����ʵ��ֵ":"1316",
		  "�͵�������":"1317","�¶ȿ��˽���(�����ս�)":"1320","�¶ȿ��˽���(�������ս�)":"1321",
		  "��������":"1322","ҵ�����𣨲������ս���":"1323","ҵ�����𣨼����ս���":"1324","�¶ȸ������Ž���":"1325",
		  "���¶Ƚ���":"1326","����1":"1327","����2":"1328","��н����":"2101","���ڲ���":"2102",
		  "���ٲ���":"2103","˰ǰ��������":"2108","ס������":"2111","�籣���˲���":"2112","��������˲���":"2113",
		  "������":"2201","˰����������":"2203","�����籣":"2204","���۹�����":"2205","˰ǰ�����ۿ�":"2301",
		  "���ڿۿ�":"2302","���ٿۿ�":"2303","ס�޿ۿ�":"2401","����ۿ�":"2402","ʧ���ۿ�":"2403",
		  "������":"2404","˰�������ۿ�":"2405","ס��ˮ��Ѽ�����ˮ�ۿ�":"2410","ס�޷ѿۿ�-�ܲ�����ס�޼�ˮ���":"2411",
		  "ס�޷ѿۿ�-�ܲ���������ˮ��":"2412","ס�޷ѿۿ�(÷��Է)":"2413","��Ů����":"/4J1",
		  "ס������":"/4J4","ס�����":"/4J5","��������":"/4J6","��������":"/4J2","ҵ������(�������ս�)":"1323",
		  "ҵ������(�����ս�)":"1324","���":"1324","ʧ��":"2403","����˰":"2405"}
df_or = DataFrame(Series(dir_or),columns=['������'])

dir_jc = {"���Ͻ��ɽ�����":"901","���Ͻ��ɽ�˾":"902","ҽ�ƽ��ɽ�����":"903",
		  "ҽ�ƽ��ɽ�˾":"904","ʧҵ���ɽ�����":"905","ʧҵ���ɽ�˾":"906",
		  "���˽��ɽ�˾":"907","�������ɽ�˾":"908","��������ɽ�����":"909",
		  "�ۺϱ��չ�˾���ɽ��":"912","Ӫҵ˰���Ͼ���":"940","�ۺϱ��ո��˽��ɽ��":"952",
		  "��������ɽ�˾":"973","������͹��ʱ�׼":"1006","פ�����":"1201",
		  "����":"1202","������":"1212","ְ������":"1214","��λ����":"1215",
		  "���ս�����":"1312","��װ��":"1318","�����׼":"1329","������":"2404",
		  "�����Ը�����˰��׼":"2800","˰��������":"2901"}
df_jc = DataFrame(Series(dir_jc),columns=['������'])

dir_kq = {"���¼�н����":"3102","Ӧ��������":"3101","ʵ�ʳ�������":"3104","��������":"3203",
		  "��н�¼�����":"3202","��������":"3204","�������":"3206","ɥ������":"3207",
		  "��������":"3201","��������":"3209","��������":"3210","���������":"3205",
		  "���˼�����":"3208","����δ�򿨣��Σ�":"3300","�ǵ���δ�򿨣��Σ�":"3301",
		  "�ǵ��̳ٵ�0-30M":"3302","�ǵ��̳ٵ�31-60M":"3303","�ǵ��̳ٵ�61-120M":"3304",
		  "�ǵ��̳ٵ�120M����":"3305","��������30M����":"3317","��������30M����":"3306",
		  "���̳ٵ�0-10M":"3307","���̳ٵ�11-30M":"3308","���̳ٵ�31-60M":"3309",
		  "���̳ٵ�61-120M":"3310","���̳ٵ�120M����":"3311","��������1Сʱ��":"3312",
		  "��������1Сʱ����":"3313","ƽʱ�Ӱ�ʱ":"3401","���ռӰ�ʱ":"3403","��ĩ�Ӱ�ʱ":"3402"}
df_kq = DataFrame(Series(dir_kq),columns=['������'])


#�Զ��庯��
def opt(df, name, form='.xls'):
	df_800 = df.loc[df.loc[:,'ϵͳ'] == 800,:]
	df_830 = df.loc[df.loc[:,'ϵͳ'] == 830,:]
	del df_800['ϵͳ']
	del df_830['ϵͳ']
	if len(df_800) >= 1:
		df_800.sort_values(by=['SAP���'])
	if len(df_830) >= 1:
		df_830.sort_values(by=['SAP���'])
	if form == '.xls':
		if len(df_800) >= 1:
			df_800.to_excel(path_char + name + "800" + form,index=False)
			print(name + "800�ѵ���,����Ϊ: " + str(len(df_800)))
		if len(df_830) >= 1:
			df_830.to_excel(path_char + name + "830" + form,index=False)
			print(name + "830�ѵ���,����Ϊ: " + str(len(df_830)))
	elif form == '.txt':
		if len(df_800) >= 1:
			df_800.to_csv(path_text + name + "800" + form,index=False,sep='\t')
			print(name + "800�ѵ���,����Ϊ: " + str(len(df_800)))
		if len(df_830) >= 1:
			df_830.to_csv(path_text + name + "830" + form,index=False,sep='\t')
			print(name + "830�ѵ���,����Ϊ: " + str(len(df_830)))
	else:
		print("δ��ָ����ʽ�ύ����!")


#����
if os.path.exists(path_mid + '����.xlsx'):
	df_atd = pd.read_excel(path_mid + '����.xlsx')
	df_atd = df_atd.melt(id_vars=['SAP���','����'], var_name="����", value_name="ʱ��")
	df_atd = pd.merge(df_atd, df_kq, left_on='����', right_index=True, how='left')
	df_atd = pd.merge(df_atd, df_date, left_on='SAP���', right_on='SAP��Ա���', how='left')
	df_atd.loc[:,'ʱ��'] = df_atd.loc[:,'ʱ��'].astype('float')
	df_attendance = DataFrame(df_atd[df_atd['ʱ��'] > 0],columns=['SAP���','����','������','��ʼ����','ʱ��','���','��λ','���','ϵͳ'])
	
	opt(df_attendance,"����")
else:
	print("δ���ֿ�������!")


#������ϸ
if os.path.exists(path_mid + '������ϸ.xlsx'):
	df_jt = pd.read_excel(path_mid + '������ϸ.xlsx')
	df_jt.rename(columns={'������':'����'},inplace=True)
	df_jt.loc[:,'���'] = df_jt.loc[:,'���'].apply(lambda x:round(x,2))
	df_jt = pd.pivot_table(df_jt,index=['SAP���','����','����'], values=['���'],aggfunc='sum').reset_index()
	df_jt.loc[df_jt['����'].str.contains("��"),"���"] = df_jt.loc[df_jt['����'].str.contains("��"),"���"].apply(lambda x:-abs(x))
else:
	df_jt = DataFrame()


#�籣
if os.path.exists(path_mid + '�籣.xlsx'):
	df_ss = pd.read_excel(path_mid + '�籣.xlsx')
	df_ss = pd.merge(df_ss,df_date,left_on='SAP���',right_on='SAP��Ա���',how='left')
	
	df_kg = DataFrame(df_ss[(df_ss['�籣�˻�']!=0)|(df_ss['�������˻�']!=0)],columns=['SAP���','����'])
	df_kg['0001'] = 'ZM'
	df_kg['0002'] = 'ZM'
	df_kg['0003'] = 'ZM'
	df_kg['0004'] = 'ZM'
	df_kg['0005'] = 'ZM'
	df_kg = df_kg.melt(id_vars=['SAP���','����'],var_name="����Ϣ����",value_name='��̯��Χ')
	df_kg = pd.merge(df_kg, df_ss, on=['SAP���','����'],how='left')
	
	df_kg.loc[:,'��̯��׼'] = "01"
	df_kg.loc[(df_kg.loc[:,'����Ϣ����']=='0001')&(df_kg.loc[:,'���Ͻ��ɽ�����']==0)&(df_kg.loc[:,'���Ͻ��ɽ�˾']!=0),"��̯��׼"] = '02'
	df_kg.loc[(df_kg.loc[:,'����Ϣ����']=='0002')&(df_kg.loc[:,'ʧҵ���ɽ�����']==0)&(df_kg.loc[:,'ʧҵ���ɽ�˾']!=0),"��̯��׼"] = '02'
	df_kg.loc[(df_kg.loc[:,'����Ϣ����']=='0003')&(df_kg.loc[:,'ҽ�ƽ��ɽ�����']==0)&(df_kg.loc[:,'ҽ�ƽ��ɽ�˾']!=0),"��̯��׼"] = '02'
	df_kg.loc[(df_kg.loc[:,'����Ϣ����']=='0003')&(df_kg.loc[:,'ҽ�ƽ��ɽ�����']==0)&(df_kg.loc[:,'ҽ�ƽ��ɽ�˾']!=0),"��̯��׼"] = '02'

	df_kg.loc[:,'�����͹�Ա֧���ķ�̯'] = ""
	df_kg.loc[(df_kg.loc[:,'����Ϣ����']=='0001')&(df_kg.loc[:,'���Ͻ��ɽ�����'] + df_kg.loc[:,'���Ͻ��ɽ�˾'] > 0),"�����͹�Ա֧���ķ�̯"] = "X"
	df_kg.loc[(df_kg.loc[:,'����Ϣ����']=='0002')&(df_kg.loc[:,'ʧҵ���ɽ�����'] + df_kg.loc[:,'ʧҵ���ɽ�˾'] > 0),"�����͹�Ա֧���ķ�̯"] = "X"
	df_kg.loc[(df_kg.loc[:,'����Ϣ����']=='0003')&(df_kg.loc[:,'ҽ�ƽ��ɽ�����'] + df_kg.loc[:,'ҽ�ƽ��ɽ�˾'] > 0),"�����͹�Ա֧���ķ�̯"] = "X"
	df_kg.loc[(df_kg.loc[:,'����Ϣ����']=='0004')&(df_kg.loc[:,'���˽��ɽ�˾'] > 0),"�����͹�Ա֧���ķ�̯"] = "X"
	df_kg.loc[(df_kg.loc[:,'����Ϣ����']=='0005')&(df_kg.loc[:,'�������ɽ�˾'] > 0),"�����͹�Ա֧���ķ�̯"] = "X"
	
	df_kg.loc[:,'�޷�̯'] = "X"
	df_kg.loc[df_kg.loc[:,'�����͹�Ա֧���ķ�̯'] == "X", "�޷�̯"] = ""
	df_kg.loc[:,'��̯����ԭ�����'] = '03'
	df_kg.loc[df_kg.loc[:,'�޷�̯'] == "X", '��̯����ԭ�����'] = '14'
	df_kg.loc[:,'��ҵ'] = '01'
	df_kg.loc[:,'��̯��'] = "ZM01"
	
	df_kg_end = DataFrame(df_kg,columns=['SAP���', '����Ϣ����','��ʼ����','��̯����ԭ�����','���յ����','��̯��Χ','��ҵ','��̯��','��̯��׼','�����͹�Ա֧���ķ�̯','�������ķ�̯֧��','�޷�̯','ϵͳ'])
	opt(df_kg_end, "�籣��̯", '.txt')
	
	
	df_kg_gjj = DataFrame(df_ss[df_ss['�������˻�'] != 0])
	df_kg_gjj.loc[:,'����'] = "ZM"
	df_kg_gjj.loc[:,'��'] = "ZM01"
	df_kg_gjj.loc[:,'����'] = "01"

	df_kg_gjj.loc[:,'�����͹�Ա����'] = ""
	df_kg_gjj.loc[df_kg_gjj.loc[:,'��������ɽ�����'] + df_kg_gjj.loc[:,'��������ɽ�˾'] > 0,"�����͹�Ա����"] = "X"
	df_kg_gjj.loc[:,'������'] = "X"
	df_kg_gjj.loc[df_kg_gjj.loc[:,'�����͹�Ա����'] == "X","������"] = ""
	df_kg_gjj_end = DataFrame(df_kg_gjj, columns=['SAP���','��ʼ����','ס���������˺�','����','��','����','�����͹�Ա����','��λ����','Ա������','������','ϵͳ'])
	opt(df_kg_gjj_end, "������", '.txt')
	
	
	df_ss_mid = pd.read_excel(path_mid + '�籣.xlsx')
	del df_ss_mid['�籣�˻�']
	del df_ss_mid['�������˻�']
	df_ss_mid = df_ss_mid.melt(id_vars=['SAP���','����'],var_name='����',value_name='���')
else:
	df_ss_mid = DataFrame()


#����ר��
if os.path.exists(path_mid + '����ר��.xlsx'):
	df_fj = DataFrame(pd.read_excel(path_mid + '����ר��.xlsx'),columns=['SAP���','����','��Ů����','ס�����','ס������','��������','��������'])
	df_fj = df_fj.melt(id_vars=['SAP���','����'], var_name="����", value_name="���")
else:
	df_fj = DataFrame()


#н���춯��
if os.path.exists(path_mid + 'н���춯��.xlsx'):
	df_xz = pd.read_excel(path_mid + 'н���춯��.xlsx')
	df_xz.loc[:,'����ԭ��'] = '50'
	df_xz.loc[:,'���ʵȼ�����'] = "01"
	df_xz.loc[:,'����'] = '0001'
	df_xz.loc[:,'����'] = 'A0'
	df_xz.loc[:,'�������ʹ�����'] = 1001
	df_xz.loc[:,'�������ʽ��'] = df_xz.loc[:,'������͹��ʱ�׼']
	df_xz.loc[:,'ְ�����ʹ�����'] = 1002
	df_xz.loc[:,'ְ�����ʽ��'] = df_xz.loc[:,'н��'] - df_xz.loc[:,'������͹��ʱ�׼']
	df_xz.loc[:,'�̶����ʱ�׼������'] = 1003
	df_xz.loc[:,'�̶�����'] = df_xz.loc[:,'н��']
	df_xz = pd.merge(df_xz, df_date,left_on='SAP���',right_on='SAP��Ա���',how='left')
	df_xz_end = DataFrame(df_xz[df_xz['н��']>0],columns = ['SAP���','����','��ʼ����','����ԭ��','���ʵȼ�����','��Χ','����','����','�������ʹ�����',\
						  '�������ʽ��','ְ�����ʹ�����','ְ�����ʽ��','�̶����ʱ�׼������','�̶�����','Сʱ��������Ŀ','ϵͳ'])
	opt(df_xz_end,"н��",'.xls')
	
	
	df_xz_dj = DataFrame(df_xz[df_xz['н��']==0],columns=['SAP��Ա���','��ʼ����','ϵͳ'])
	opt(df_xz_dj,"н�ʶ���",'.txt')
	
	
	df_xz_jc = DataFrame(df_xz, columns=['SAP���','����','������͹��ʱ�׼'])
	df_xz_jc = df_xz_jc.melt(id_vars=['SAP���','����'], var_name='����', value_name='���')
else:
	df_xz_jc = DataFrame()


#�����춯��
if os.path.exists(path_mid + '�����춯��.xlsx'):
	df_jty = pd.read_excel(path_mid + '�����춯��.xlsx')
	df_jty.rename(columns={'��Ŀ':'����'}, inplace=True)
else:
	df_jty = DataFrame()


#Сʱ��
if os.path.exists(path_mid + 'Сʱ��.xlsx'):
	df_sl = DataFrame(pd.read_excel(path_mid + 'Сʱ��.xlsx'),columns=['SAP���','����','Сʱ��','ʱн','����','��н','���','ʧ��','����˰'])
	
	df_sl.fillna(0,inplace=True)
	df_sl.loc[:,'Сʱ������'] = df_sl.loc[:,'Сʱ��'] * df_sl.loc[:,'ʱн'] + df_sl.loc[:,'����'] * df_sl.loc[:,'��н']
	df_sl = DataFrame(df_sl, columns=['SAP���','����','Сʱ������','���','ʧ��','����˰'])
	df_sl = pd.pivot_table(df_sl, index=['SAP���'], values=['Сʱ������','���','ʧ��','����˰'], aggfunc='sum').reset_index()

	df_sl_end = df_sl.melt(id_vars=['SAP���'], var_name="����", value_name='���')
	df_sl_end = DataFrame(df_sl_end, columns=['SAP���','����','����','���'])
else:
	df_sl_end = DataFrame()


#������ϸ
if os.path.exists(path_mid + '������ϸ.xlsx'):
	df_bk = DataFrame(pd.read_excel(path_mid + '������ϸ.xlsx'))
	df_bk.loc[:,'��Ҫ����'] = 0
	df_bk = pd.merge(df_bk, df_date, left_on='SAP���', right_on='SAP��Ա���', how='left')
	df_bk_end = DataFrame(df_bk, columns=['SAP���','��ʼ����','��������','��Ҫ����','����','���д���','�����˺�','ϵͳ'])
	opt(df_bk_end, "������ϸ",'.txt')
else:
	print("δ��������������Ϣ!")


#����������
df_jcx = pd.concat([df_ss_mid,df_jty,df_xz_jc], axis=0)
df_jcx = pd.merge(df_jcx, df_date, left_on='SAP���',right_on='SAP��Ա���',how='left')
df_jcx = pd.merge(df_jcx, df_jc, left_on='����', right_index=True,how='left')
df_jcx = df_jcx[df_jcx['������'].notnull()]
df_jcx.loc[:,"������"] = df_jcx.loc[:,"������"].astype('int').astype('str')
df_jcx.loc[:,"������"] = df_jcx.loc[:,"������"].apply(lambda x:x.zfill(4))

df_jcx_opt = DataFrame(df_jcx[(df_jcx['���']>0)&(df_jcx['������'].notnull())],columns=['SAP���','����','��ʼ����','��������','������','���','���','��λ','ϵͳ'])
opt(df_jcx_opt, "������", '.xls')

df_jcx_dj = DataFrame(df_jcx[(df_jcx['���']==0)&(df_jcx['������'].notnull())],columns=['SAP���','������','��ʼ����','ϵͳ'])
df_jcx_dj = DataFrame(df_jcx_dj[(df_jcx_dj['������']=="1214")|(df_jcx_dj['������']=="1202")|(df_jcx_dj['������']=="1201")|(df_jcx_dj['������']=="1212")|(df_jcx_dj['������']=="1215")])
opt(df_jcx_dj, "�����Զ���", '.txt')


#żȻ������
df_orx = pd.concat([df_jt,df_ss_mid,df_fj,df_sl_end],axis=0)
df_orx = pd.merge(df_orx, df_date, left_on='SAP���',right_on='SAP��Ա���',how='left')
df_orx = pd.merge(df_orx, df_or, left_on='����', right_index=True,how='left')
df_orx = df_orx[df_orx['������'].notnull()]
df_orx.loc[:,'������'] = df_orx.loc[:,'������'].astype('str')
df_orx.loc[df_orx['������'].str.find(".")!=-1,"������"] = df_orx.loc[df_orx['������'].str.find(".")!=-1,"������"].astype('int').astype('str')
df_orx.loc[:,"������"] = df_orx.loc[:,"������"].apply(lambda x:x.zfill(4))

df_orx_opt = DataFrame(df_orx[(df_orx['���']>0)&(df_orx['������'].notnull())&(df_orx['��ʼ����'].notnull())],columns=['SAP���','����','������','���','���','��λ','��ʼ����','������','ϵͳ'])
df_orx_opt.loc[:,'��ʼ����'] = df_orx_opt.loc[:,'��ʼ����'].astype('int')
df_orx_opt.loc[:,'ϵͳ'] = df_orx_opt.loc[:,'ϵͳ'].astype('int')
opt(df_orx_opt,"żȻ��", '.xls')

print("\nģ����Ѿ��������,�뵼��ϵͳ!")
input()
