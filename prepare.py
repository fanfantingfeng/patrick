#--coding:utf-8--

import pandas as pd
import numpy as np
from pandas import Series, DataFrame
import datetime
import arrow
from dateutil.relativedelta import relativedelta
import calendar

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
data = pd.read_excel(root_file + "\\" + str(last_m.month) + "月人事异动表-800.xlsx",engine='openpyxl')
data.loc[:,"开始日期"] = startday
data.loc[data['入职'] == "新员工入职", "开始日期"] = data.loc[data['入职'] == "新员工入职", "入职日期"].apply(lambda x:x.strftime("%Y%m%d"))
data.loc[:,"结束日期"] = "99991231"

file = 'D:\\根目录\\参数表\\参数表.xlsx'
city = pd.read_excel(file,sheet_name='人事范围',engine='openpyxl')
cost_center = pd.read_excel(file,sheet_name='成本中心',engine='openpyxl',dtype={'特殊成本中心':str})

data = pd.merge(data, city, left_on="人事范围描述", right_on="人事范围", how="left")


#特殊的人员过账成本中心
dir_special = {4002698:4800200300, 4004623:4800200300}
df_special = pd.DataFrame(pd.Series(dir_special), columns=['过账成本中心'])
#df_special['人员编号'] = df_special.index


#自定义函数
def output(data,text):
	if len(data) >= 1:
		data.to_csv(work_file + "\\" + text + ".txt", sep='\t', index=False)
		print(text + "已导出,条数为: " + str(len(data)))
	else:
		print("无需导出" + text)

#获取某日期对应月份次月的最后一天
def get_endofmonth_next(date):
	date = date + relativedelta(months=1)
	weekday,last_num = calendar.monthrange(date.year, date.month)
	date_new = datetime.datetime(date.year,date.month,day=last_num)
	return (date_new).strftime("%Y%m%d")


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
	df_cc.loc[:,"有无事件"] = "无"
	
	df_cc.loc[df_cc['入职'].notnull(),"有无事件"] = "有"
	df_cc.loc[df_cc['调动'].notnull(),"有无事件"] = "有"
	df_cc.loc[df_cc['晋级'].notnull(),"有无事件"] = "有"
	df_cc.loc[df_cc['特殊事件'].notnull(),"有无事件"] = "有"
	df_cc[['总部管理中心/事业部','中心','分中心','部门/零售管理公司','分部门','组/分公司部门','店铺/专柜','组织单元描述']] = df_cc[['总部管理中心/事业部','中心','分中心','部门/零售管理公司','分部门','组/分公司部门','店铺/专柜','组织单元描述']].fillna("无")
	list3 = [df_cc['中心'],df_cc['分中心'],df_cc['部门/零售管理公司'],df_cc['分部门'],df_cc['组/分公司部门'],df_cc['店铺/专柜'],df_cc['组织单元描述']]
	df_cc.loc[:,'组织参数'] = df_cc['总部管理中心/事业部'].str.cat(list3)

	df_cc['品牌'] = "其他"
	df_cc.loc[df_cc['组织参数'].str.contains('财务'), '品牌'] = '总直'
	df_cc.loc[df_cc['组织参数'].str.contains('人力资源'), '品牌'] = '总直'
	df_cc.loc[df_cc['组织参数'].str.contains('bonwe品牌'), '品牌'] = 'MB'
	df_cc.loc[df_cc['组织参数'].str.contains('CITY品牌'), '品牌'] = 'MC'
	df_cc.loc[df_cc['组织参数'].str.contains('童装'), '品牌'] = '童装'
	df_cc.loc[df_cc['组织参数'].str.contains('Moomoo品牌'), '品牌'] = 'MM'
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
	df_cc = pd.merge(df_cc, df_special, left_on='SAP人员编号', right_index=True,how='left')
	df_cc.loc[:,'成本中心'] = "无"
	
	df_cc.loc[df_cc['特殊成本中心'].notnull(),'成本中心'] = df_cc.loc[df_cc['特殊成本中心'].notnull(),'特殊成本中心']
	df_cc.loc[df_cc['组织成本中心'].notnull(),'成本中心'] = df_cc.loc[df_cc['组织成本中心'].notnull(),'组织成本中心']
	df_cc.loc[df_cc['过账成本中心'].notnull(),'成本中心'] = df_cc.loc[df_cc['过账成本中心'].notnull(),'过账成本中心']
	df_cc.loc[df_cc['成本中心']!='无', "成本中心"] = df_cc.loc[df_cc['成本中心']!='无', "成本中心"].apply(lambda x:int(x))
	df_cc.loc[:,'是否调整成本中心'] = "否"
	df_cc.loc[df_cc['人员成本中心'] != df_cc['成本中心'], "是否调整成本中心"] = "是"
	df_cc.loc[df_cc['有无事件'] == '有', "是否调整成本中心"] = "是"
	df_cc = DataFrame(df_cc[(df_cc['是否调整成本中心'] == "是")])

	df_cc.loc[:,'订单_1'] = ""
	df_cc.loc[df_cc['人事范围描述'].str.contains("MB"), '订单_1'] = "10"
	df_cc.loc[df_cc['人事范围描述'].str.contains("MC"), '订单_1'] = "20"
	df_cc.loc[:,'公司代码'] = df_cc['成本中心'].apply(lambda x:str(x)[:4])
	df_cc.loc[:,'订单_3'] = "0102"
	fil_dz = (df_cc['店铺职级'].str.contains("店长"))|(df_cc['店铺职级'].str.contains("店助"))|(df_cc['店铺职级'].str.contains("店铺形象"))|(df_cc['店铺职级'].str.contains("商品"))
	df_cc.loc[fil_dz,"订单_3"] = "0103"
	fil_dg = (df_cc['店铺职级'].str.contains("导购"))|(df_cc['店铺职级'].str.contains("店员"))|(df_cc['店铺职级'].str.contains("训练员"))|(df_cc['店铺职级'].str.contains("牛人"))
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
df.reset_index()

column_incident = ["入职", "转正", "转正原因", "调动", "调动原因", "晋级", "晋/降级", "离职", "离职原因", "特殊事件", "特殊事件原因","人事范围描述.1"]
df.loc[:, column_incident] = df.loc[:, column_incident].fillna("无")

df.loc[:,"有无事件"] = "有"
df.loc[(df['入职'] == "无")&(df['转正'] == "无")&(df['调动'] == "无")&(df['晋级'] == "无")&(df['离职'] == "无")&(df['特殊事件'] == "无"),"有无事件"] = "无"
df.loc[(df['调动原因']=="公司内调动"), "有无事件"] = "无"
#df.loc[(df['离职']=="无"), "有无事件"] = "无"?

df.loc[:,"是否更换法人公司"] = "否"
df.loc[(df['有无事件'] == "有")&(df['人事范围描述.1'] != "无")&(df.loc[:,'人事范围描述'].apply(lambda x:x[:2])!=df.loc[:,'人事范围描述.1'].apply(lambda x:x[:2])),"是否更换法人公司"] = "是"
df.loc[df['特殊事件原因']=="合同改签","是否更换法人公司"] = "是"

df.loc[:,"是否隔月重入职"] = "否"
df.loc[:,"上一次离职日期"] = df.loc[:,"上一次离职日期"].fillna(datetime.datetime(1900,1,1))
#df.loc[:,"上一次离职日期"] = df.loc[:,"上一次离职日期"].apply(lambda x:datetime.datetime.strptime(x,"%Y-%m-%d"))
#df.loc[:,"入职日期"] = df.loc[:,"入职日期"].apply(lambda x:datetime.datetime.strptime(x,"%Y-%m-%d"))
df.loc[:,'核对日期1'] = df.loc[:,'上一次离职日期'].apply(lambda x:datetime.datetime((x+relativedelta(months=1)).year,(x+relativedelta(months=1)).month,1))
df.loc[:,'核对日期2'] = df.loc[:,'入职日期'].apply(lambda x:datetime.datetime(x.year,x.month,1))
df.loc[(df['核对日期1'] == df['核对日期2'])&(df['上一次离职日期'].apply(lambda x:x.day == 1))&(df['入职']!="无"), "是否隔月重入职"] = "是"
df.loc[(df['核对日期1'] < df['核对日期2'])&(df['上一次离职日期'] != datetime.datetime(1900,1,1))&(df['入职']!="无") ,"是否隔月重入职"] = "是"

df.loc[:,"纳税终止日期"] = ""
df.loc[df['离职'] != "无", "纳税终止日期"] = df.loc[df['离职'] != "无", "离职日期"].apply(lambda x:get_endofmonth_next(x))

df.loc[:,"是否重置累计"] = ""
df.loc[df['是否更换法人公司'] == "是", "是否重置累计"] = "X"
df.loc[(df['是否隔月重入职'] == "是")&(df['入职'] == "重新入职"), "是否重置累计"] = "X"
df.loc[(df['转正原因'].str.contains("实习"))|(df['特殊事件原因'].str.contains("临时"))|(df['特殊事件原因'].str.contains("派遣"))|(df['特殊事件原因'].str.contains("实习"))|(df['特殊事件原因'].str.contains("改签")),"是否重置累计"] = "X"

df.loc[:,"开始日期"] = df['年'].astype('str') + "-" + df['月'].astype('str') + "-" + str(1)
df.loc[:,"开始日期"] = df['开始日期'].apply(lambda x:datetime.datetime.strptime(x,"%Y-%m-%d").strftime("%Y%m%d"))
df.loc[df['入职'] == "新员工入职", "开始日期"] = df.loc[df['入职'] == "新员工入职", "入职日期.1"].apply(lambda x:x.strftime("%Y%m%d"))

df.loc[:,'判定日期'] = ""
df.loc[(df['入职'] != "无")|(df['是否更换法人公司'] == "是")|((df['是否隔月重入职'] == "是")&(df['入职'] != "无"))|(df['是否重置累计'] == 'X'), "判定日期"] = df.loc[(df['入职'] != "无")|(df['是否更换法人公司'] == "是")|(df['是否隔月重入职'] == "是")|(df['是否重置累计'] == 'X'), "开始日期"]
df.loc[df['最近一次税收录入日期'].notnull(), "判定日期"] = df.loc[df['最近一次税收录入日期'].notnull(), "最近一次税收录入日期"].apply(lambda x:x.strftime("%Y%m%d"))

df.loc[:,"税收录入日期"] = ""
df.loc[df['是否重置累计'] == "X", "税收录入日期"] = df.loc[df['是否重置累计'] == "X", "判定日期"]
df.loc[df['是否更换法人公司'] == "是", "税收录入日期"] = df.loc[df['是否更换法人公司'] == "是", "判定日期"]
df.loc[(df['是否隔月重入职'] == "是")&(df['入职']!="无"), "税收录入日期"] = df.loc[(df['是否隔月重入职']=="是")&(df['入职']!="无"), "开始日期"]
df.loc[df['最近一次税收录入日期'].notnull(), "税收录入日期"] = df.loc[df['最近一次税收录入日期'].notnull(), "最近一次税收录入日期"].apply(lambda x:x.strftime("%Y%m%d"))
df.loc[df['入职'] != "无", "税收录入日期"] = df.loc[df['入职'] == "新员工入职", "开始日期"]

df.loc[:,'税收类型'] = "0"
df.loc[(df['员工组'].str.contains("临时员工"))|(df['员工组'].str.contains("实习生")), '税收类型'] = "4"
df.loc[df['员工组'].str.contains("外籍"), "税收类型"] = "2"

df_tax = DataFrame(df[df['有无事件']=="有"], columns=["SAP人员编号", "开始日期", "结束日期", "征税地区", "税组", "税收类型", "税收录入日期", "纳税终止日期", "是否免税", "是否重置累计"])
if len(df_tax)>=1:
	output(df_tax,"7-个税")
	print('''Project:    MB_HR_PERSON
Subproject: 0531
Object:     0531_01
''')
else:
	print("无需导出个税信息!")

print("模板表已制作完成,请导入SAP系统!")
input()
