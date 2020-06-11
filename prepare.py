#--coding:utf-8--

import pandas as pd
import numpy as np
from pandas import Series, DataFrame
import datetime
import arrow
from dateutil.relativedelta import relativedelta

print('''
本模板表适用于当月人事异动表的准备工作
作者: 凡凡
日期: 2020/6/10
''')

#指定文件夹路径
root_file = 'D:\\根目录\\人事异动'
work_file = 'C:\\Users\\hq01ug601\\Desktop\\1-text'


#信息准备
now = datetime.datetime.now()
last_m = now + relativedelta(months=-1)
startday = datetime.datetime(last_m.year, last_m.month, 1).strftime("%Y%m%d")
endday = (datetime.datetime(now.year,now.month,1) - datetime.timedelta(1,0,0,0)).strftime("%Y%m%d")


#数据准备
data = pd.read_excel(root_file + "\\" + str(last_m.month) + "月人事异动表-800.xlsx")
data.loc[:,"开始日期"] = startday
data.loc[data['入职'] == "新员工入职", "开始日期"] = data.loc[data['入职'] == "新员工入职", "入职日期"].apply(lambda x:x.strftime("%Y%m%d"))
data.loc[:,"结束日期"] = "99991231"

file = 'D:\\根目录\\参数表\\参数表.xlsx'
city = pd.read_excel(file,sheet_name='Sheet1')
cost_center = pd.read_excel(file,sheet_name='Sheet2',dtype={'特殊成本中心':str})

data = pd.merge(data, city, left_on="人事范围描述", right_on="人事范围", how="left")


#自定义函数
def output(data,text):
	if len(data) >= 1:
		data.to_csv(work_file + "\\" + text + ".txt", sep='\t', index=False)
		print(text + "已导出,条数为: " + str(len(data)))
	else:
		print("无需导出" + text)



#计划工作时间
df_plan = DataFrame(data[data['入职'].notnull()], columns=["SAP人员编号", "开始日期"])
output(df_plan,"2-计划工作时间")
print('''Project:    PT-PY-兰州
Subproject: PT-PY-WWF
Object:     计划工作时间
''')


#社保无分摊
df_ss = DataFrame(data[data['入职'].notnull()], columns=["SAP人员编号", "开始日期"])
list1 = ['0001', '0002', '0003', '0004', '0005']
for i in list1:
	df_ss[i] = "01"
df_ss = df_ss.melt(id_vars=['SAP人员编号', '开始日期'], value_vars=list1, var_name='子信息类型', value_name='分摊标准')
if len(df_ss)>=1:
	df_ss.loc[:,'分摊更改原因代码'] = ""
	df_ss.loc[:,'保险单编号'] = ""
	df_ss.loc[:,'行业'] ="01"
	df_ss.loc[:,'组(ZM01)'] = 'ZM01'
	df_ss.loc[:,'分摊范围'] = "ZM"
	df_ss.loc[:,'雇主和雇员支付的分摊'] = ""
	df_ss.loc[:,'仅雇主的分摊支付'] = ""
	df_ss.loc[:,'无分摊'] = 'X'
	df_ss.loc[:,"开始日期"] = startday
	order_ss = ['SAP人员编号','子信息类型','开始日期','分摊更改原因代码','保险单编号',
			'分摊范围','行业','组(ZM01)','分摊标准',
			'雇主和雇员支付的分摊','仅雇主的分摊支付','无分摊']
	df_ss = df_ss[order_ss]
	output(df_ss, "1-社保无分摊")
	print('''Project:    PT-PY-兰州
Subproject: PT-PY-WWF
Object:     0007-社保
''')
else:
	print("无需导出社保数据!")


#公积金无分摊
df_fund = DataFrame(data[data['入职'].notnull()], columns=["SAP人员编号", "开始日期"])
if len(df_fund)>=1:
	df_fund.loc[:,'地区(ZM)'] = 'ZM'
	df_fund.loc[:,'组(ZM01)'] = 'ZM01'
	df_fund.loc[:,'级别(01)'] = '01'
	df_fund.loc[:,'无分摊'] = "X"
	df_fund.loc[:,'开始日期'] = startday
	order_fund = ['SAP人员编号','开始日期','住房公积金账号','地区(ZM)','组(ZM01)','级别(01)','雇主和雇员支付的分摊','雇主支付','雇员支付','无分摊']

	df_fund = DataFrame(df_fund, columns=order_fund)
	output(df_fund, "4-公积金无分摊")
	print('''Project:    PT-PY-兰州
Subproject: PT-PY-WWF
Object:     GONGJJJJ
''')
else:
	print("无需导出公积金数据!")


#现金
df_cash = DataFrame(data[data['入职'].notnull()],columns=['SAP人员编号', '开始日期', '结束日期', '分公司'])
if len(df_cash)>=1:
	output(df_cash, "3-现金")
	print('''Project:    PT-PY-兰州
Subproject: PT-PY-WWF
Object:     0010-无银行卡
''')
else:
	print("无需导出现金数据!")


#经常性定界
df_stp = DataFrame(data[data['入职'] == '重新入职'], columns=['SAP人员编号', '开始日期'])
if len(df_stp)>=1:
	list2 = ['1006','1201','1202','1212','1214','1215','2404','2800','2901']
	for i in list2:
		df_stp[i] = np.nan
	df_stp = df_stp.melt(id_vars=['SAP人员编号', '开始日期'], value_vars=list2, var_name='子信息类型', value_name='无效值')
	df_stp = DataFrame(df_stp, columns=['SAP人员编号', '子信息类型', '开始日期'])
	output(df_stp, '5-经常性定界')
	print('''Project:    PT-PY
Subproject: PT-PY-FY
Object:     0014定界
''')
else:
	print("无需导出经常性定界信息!")


#成本中心
df_cc = DataFrame(data[(data['员工组'] != "临时员工")&(~data['人事范围描述'].str.contains("总部"))])
if len(df_cc)>=1:
	df_cc.loc[:,"有无事件"] = "有"
	fil = (df_cc['入职'].isnull())&(df_cc['调动'].isnull())&(df_cc['晋级'].isnull())&(df_cc['特殊事件'].isnull())
	df_cc.loc[fil,"有无事件"] = "无"
	df_cc[['总部管理中心/事业部','中心','分中心','部门/零售管理公司','分部门','组/分公司部门','店铺/专柜','组织单元描述']] = df_cc[['总部管理中心/事业部','中心','分中心','部门/零售管理公司','分部门','组/分公司部门','店铺/专柜','组织单元描述']].fillna("无")
	list3 = [df_cc['中心'],df_cc['分中心'],df_cc['部门/零售管理公司'],df_cc['分部门'],df_cc['组/分公司部门'],df_cc['店铺/专柜'],df_cc['组织单元描述']]
	df_cc.loc[:,'组织参数'] = df_cc['总部管理中心/事业部'].str.cat(list3)

	df_cc['品牌'] = "其他"
	df_cc.loc[df_cc['组织参数'].str.contains('财务'), '品牌'] = '总直'
	df_cc.loc[df_cc['组织参数'].str.contains('人力资源'), '品牌'] = '总直'
	df_cc.loc[df_cc['组织参数'].str.contains('bonwe品牌'), '品牌'] = 'MB'
	df_cc.loc[df_cc['组织参数'].str.contains('CITY品牌'), '品牌'] = 'MC'
	df_cc.loc[df_cc['组织参数'].str.contains('童装'), '品牌'] = '童装'
	df_cc.loc[df_cc['组织参数'].str.contains('褀'), '品牌'] = '褀'
	df_cc['职能'] = "其他"
	df_cc.loc[df_cc['组织单元描述'].str.endswith('销售公司'),'职能'] = '分公司'
	df_cc.loc[df_cc['组织单元描述'].str.endswith('公司零售管理部'),'职能'] = '零售'
	df_cc.loc[df_cc['组织单元描述'].str.endswith('直营市场'),'职能'] = '直营'
	df_cc.loc[df_cc['组织参数'].str.contains('管理培训'),'职能'] = '管培生'
	df_cc.loc[df_cc['组织参数'].str.contains('行政'),'职能'] = '行政'
	df_cc.loc[df_cc['组织参数'].str.contains('财务'),'职能'] = 'FI'
	df_cc.loc[df_cc['组织参数'].str.contains('人力资源'),'职能'] = 'HR'
	df_cc.loc[df_cc['组织参数'].str.contains('工程'),'职能'] = '工程'
	df_cc.loc[df_cc['组织参数'].str.contains('渠道开发'),'职能'] = '渠道2'
	df_cc.loc[df_cc['组织参数'].str.contains('渠道发展'),'职能'] = '渠道1'
	df_cc.loc[df_cc['组织参数'].str.contains('代理人'),'职能'] = '代理人'
	df_cc.loc[(df_cc['组织参数'].str.contains('加盟市场'))&(~df_cc['组织参数'].str.contains('店群'))&(~df_cc['组织参数'].str.contains('片区')),'职能'] = '加盟'
	df_cc.loc[(df_cc['组织参数'].str.contains('加盟管理'))&(~df_cc['组织参数'].str.contains('店群'))&(~df_cc['组织参数'].str.contains('片区')),'职能'] = '加盟'

	df_cc = pd.merge(df_cc,cost_center,left_on=['分公司','品牌','职能'],right_on=['属性2','品牌','属性'],how='left')
	df_cc.loc[:,'成本中心'] = "无"
	df_cc.loc[df_cc['特殊成本中心'].notnull(),'成本中心'] = df_cc.loc[df_cc['特殊成本中心'].notnull(),'特殊成本中心']
	df_cc.loc[df_cc['组织成本中心'].notnull(),'成本中心'] = df_cc.loc[df_cc['组织成本中心'].notnull(),'组织成本中心']
	df_cc.loc[df_cc['成本中心']!='无', "成本中心"] = df_cc.loc[df_cc['成本中心']!='无', "成本中心"].apply(lambda x:int(x))
	df_cc.loc[:,'是否调整成本中心'] = "否"
	df_cc.loc[df_cc['人员成本中心'] != df_cc['成本中心'], "是否调整成本中心"] = "是"
	df_cc = DataFrame(df_cc[(df_cc['有无事件'] == "有")])

	df_cc.loc[:,'订单_1'] = ""
	df_cc.loc[df_cc['人事范围描述'].str.contains("MB"), '订单_1'] = "10"
	df_cc.loc[df_cc['人事范围描述'].str.contains("MC"), '订单_1'] = "20"
	df_cc.loc[:,'公司代码'] = df_cc['成本中心'].apply(lambda x:str(x)[:4])
	df_cc.loc[:,'订单_3'] = "0102"
	fil_dz = (df_cc['店铺职级'].str.contains("店长"))|(df_cc['店铺职级'].str.contains("店助"))|(df_cc['店铺职级'].str.contains("店铺形象"))|(df_cc['店铺职级'].str.contains("商品"))
	df_cc.loc[fil_dz,"订单_3"] = "0103"
	fil_dg = (df_cc['店铺职级'].str.contains("导购"))|(df_cc['店铺职级'].str.contains("店员"))|(df_cc['店铺职级'].str.contains("训练员"))
	df_cc.loc[fil_dg, "订单_3"] = "0101"

	df_cc.loc[df_cc['店铺职级'].isnull(),"订单编号"] = ""
	df_cc.loc[df_cc['店铺职级'].notnull(),"订单编号"] = df_cc['订单_1'].str.cat([df_cc['公司代码'], df_cc['订单_3']])
	df_cc.loc[:,'分配'] = "01"

	df_cc = DataFrame(df_cc[df_cc['成本中心']!="无"],columns=['SAP人员编号','开始日期','结束日期','分配','公司代码','成本中心','订单编号'])
	output(df_cc, "6-成本中心")
	print('''Project:    PT-PY
Subproject: PT-PY-FY
Object:     0027-成本分配
''')
else:
	print("无需导出成本中心!")


#个人所得税
df = DataFrame(data)

fil_tax = (df['入职'].notnull())|(df['调动'].notnull())|(df['晋级'].notnull())|(df['特殊事件'].notnull())|(df['离职'].notnull())|(df['转正'].notnull())
df_tax= DataFrame(df[fil_tax])
df_tax = df_tax.reset_index()

if len(df_tax)>=1:
	df_tax.loc[:,'税组'] = df_tax.loc[:,'征税地区'] + "01"
	df_tax.loc[:,'税收类型'] = "0"
	df_tax.loc[(df_tax['员工组'].str.contains("临时员工"))|(df_tax['员工组'].str.contains("实习")),"税收类型"] = "4"
	df_tax.loc[df_tax['员工组'].str.contains("外籍"),"税收类型"] = "2"

	df_tax['纳税终止日期'] = ""
	df_tax.loc[df_tax['离职'] != "无","纳税终止日期"] = (datetime.datetime(now.year,now.month+1,1) - datetime.timedelta(1,0,0,0)).strftime("%Y%m%d")
	df_tax['是否免税'] = ""

#判定两日期是否隔月的自定义函数
	def gap(x,y):
		x = x + relativedelta(months=1)
		#1号离职(1表示隔月,0表示不隔夜)
		if x.day == 1:
			if (datetime.datetime(x.year, x.month, 1) <= datetime.datetime(y.year, y.month, 1)):
				return 1
			else:
				return 0
		else:
			if (datetime.datetime(x.year, x.month, 1) < datetime.datetime(y.year, y.month, 1)):
				return 1
			else:
				return 0

	for i in df_tax.index:
		if (pd.isna(df_tax.loc[i,'上次入职日期']))&(pd.isna(df_tax.loc[i,'最近一次税收录入日期'])):
			df_tax.loc[i,'判定日期'] = ""
		elif (pd.isna(df_tax.loc[i,'上次入职日期']))&(pd.notna(df_tax.loc[i,'最近一次税收录入日期'])):
			df_tax.loc[i,'判定日期'] = df_tax.loc[i,'最近一次税收录入日期'].strftime("%Y%m01")
		elif (pd.notna(df_tax.loc[i,'上次入职日期']))&(pd.isna(df_tax.loc[i,'最近一次税收录入日期'])):
			df_tax.loc[i,'判定日期'] = df_tax.loc[i,'上次入职日期'].strftime("%Y%m01")
		elif (df_tax.loc[i,'最近一次税收录入日期'] > df_tax.loc[i,'上次入职日期']):
			df_tax.loc[i,'判定日期'] = df_tax.loc[i,'最近一次税收录入日期'].strftime("%Y%m01")
		else:
			df_tax.loc[i,'判定日期'] = df_tax.loc[i,'上次入职日期'].strftime("%Y%m01")


	for i in range(len(df_tax)):
		if (pd.notna(df_tax.loc[i,'特殊事件原因']))&((df_tax.loc[i,'特殊事件原因']=="合同改签")|(df_tax.loc[i,'特殊事件原因']=="实习生转正式员工")|(df_tax.loc[i,'特殊事件原因']=="临时员工转正式员工")):
			df_tax.loc[i,'是否重置累计'] = "X"
		elif pd.isna(df_tax.loc[i,'上一次离职日期']):
			df_tax.loc[i,'是否重置累计'] = ""
		elif (df_tax.loc[i,"入职"] == "重新入职")&(gap(df_tax.loc[i,'上一次离职日期'], df_tax.loc[i,'入职日期'])==1):
			df_tax.loc[i,'是否重置累计'] = "X"
		else:
			df_tax.loc[i,'是否重置累计'] = ""

	df_tax.loc[:,'人事范围描述.1'].fillna("无",inplace=True)

	for i in df_tax.index:
		if df_tax.loc[i,'是否重置累计'] == "X":
			df_tax.loc[i,'税收录入日期'] = df_tax.loc[i,'开始日期']
		elif (pd.notna(df_tax.loc[i,'调动']))&(df_tax.loc[i,'人事范围描述'][:2] != df_tax.loc[i,'人事范围描述.1'][:2]):
			df_tax.loc[i,'税收录入日期'] = df_tax.loc[i,'调动日期'].strftime("%Y%m01")
		elif pd.isna(df_tax.loc[i,'上一次离职日期']):
			df_tax.loc[i,'税收录入日期'] = ""
		elif (df_tax.loc[i,'入职'] == "重新入职")&(gap(df_tax.loc[i,'上一次离职日期'], df_tax.loc[i,'入职日期'])==1):
			df_tax.loc[i,'税收录入日期'] = df_tax.loc[i,'判定日期']
		elif (pd.notna(df_tax.loc[i,'最近一次税收录入日期']))&(pd.notna(df_tax.loc[i,'离职'])):
			df_tax.loc[i,'税收录入日期'] = df_tax.loc[i,'最近一次税收录入日期'].strftime("%Y%m01")
		else:
			df_tax.loc[i,'税收录入日期'] = ""

	col_tax = ['SAP人员编号','开始日期','结束日期', '征税地区','税组','税收类型','税收录入日期','纳税终止日期','是否免税','是否重置累计']
	df_tax = DataFrame(df_tax, columns=col_tax)
	output(df_tax,"7-个税")
	print('''Project:    MB_HR_PERSON
Subproject: 0531
Object:     0531_01
''')
else:
	print("无需导出个税信息!")

input()
