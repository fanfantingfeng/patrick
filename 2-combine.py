#coding=GBK

import pandas as pd
from pandas import DataFrame
import numpy as np
import os
import xlrd

print("欢迎使用合并表格模板\n一切解释权均归开发者所有!\n开发者: 凡凡\n")

#参数指定
desk = os.path.join(os.path.expanduser("~"),"Desktop")
path = desk + "\\工资明细"
path_mid = 'D:\\根目录\\项目\\中间表格'
print(path)
sheet_dir = {"考勤":"kq","津贴":"jt","社保":"ss",
			 "附加-税务":"sw","附加-SAP":"sap",
			 "薪资异动表":"xz","津贴异动表":"jt1",
			 "小时工":"sl","银行明细":"bk"}

col_kq = ["SAP编号","姓名","当月计薪天数","应出勤天数","实际出勤天数",
			"全勤奖","饭贴","大店津贴","病假天数","无薪事假天数","产假天数",
			"婚假天数","丧假天数","空勤天数","旷工天数","出差天数","陪产假天数",
			"工伤假天数","调休小时","年休天数","有薪事假天数","店铺未打卡（次）",
			"非店铺未打卡（次）","非店铺迟到0-30M","非店铺迟到31-60M",
			"非店铺迟到61-120M","非店铺迟到120M以上","行政早退30M以内",
			"行政早退30M以上","店铺迟到0-10M","店铺迟到11-30M",
			"店铺迟到31-60M","店铺迟到61-120M","店铺迟到120M以上",
			"店铺早退1小时内","店铺早退1小时以上",
			"平时加班时","节日加班时","周末加班时"]
col_jt = ["SAP编号","姓名","工资项","金额"]

col_ss = ["SAP编号","姓名","社保账户","公积金账户",
		"养老缴纳金额个人","养老缴纳金额公司","医疗缴纳金额个人","医疗缴纳金额公司",
		"失业缴纳金额个人","失业缴纳金额公司","工伤缴纳金额公司","生育缴纳金额公司",
		"公积金缴纳金额个人","公积金缴纳金额公司","养老补缴金额个人","养老补缴金额公司",
		"医疗补缴金额个人","医疗补缴金额公司","失业补缴金额个人","失业补缴金额公司",
		"工伤补缴金额公司","生育补缴金额公司","公积金补缴金额个人","公积金补缴金额公司",
		"附加医疗缴纳金额个人","附加医疗缴纳金额公司",
		"附加医疗补缴金额个人","附加医疗补缴金额公司"]


#数据初始化
mid_file = os.listdir(path_mid)
if len(mid_file) >= 1:
	for i in mid_file:
		filename = path_mid + "\\" + i
		if os.path.isfile(filename):
			os.remove(filename)
	print("初始化已完成!")
else:
	print("数据无需初始化!")



#自定义函数
def output(data,text):
	data.dropna(axis=0,how='all',inplace=True)
	data.dropna(axis=1,how='all',inplace=True)
	data = data.drop_duplicates()
	data = data.replace(" ", 0)
	data.fillna(0,inplace=True)
	df = DataFrame(data[data.loc[:,'SAP编号'].notnull()])
	try:
		df.loc[:,'SAP编号'] = df.loc[:,'SAP编号'].astype('int')
	except ValueError:
		df.loc[:,'SAP编号'] = df.loc[:,'SAP编号'].astype('str')
	if len(df.index) >= 1:
		df.loc[:,'SAP编号'] = df.loc[:,'SAP编号'].astype('str')
		df = df[(df.loc[:,'SAP编号'].notnull())&(df.loc[:,'SAP编号'].str.isnumeric())]
		df = df.fillna(0)
		df.loc[:,'SAP编号'] = df.loc[:,'SAP编号'].astype('int')
		filename = path_mid + "\\" + text + '.xlsx'
		if len(df.index) >= 1:
			df.to_excel(filename, index=False)
			print(text + "表已生成,容量为: " + str(len(df)))
		else:
			print(text + "数据无需导出!")
	else:
		print(text + "数据未发现!")

def stan(data):
	data.dropna(axis=0,how='all',inplace=True)
	data.dropna(axis=1,how='all',inplace=True)
	data.loc[:,'SAP编号'] = data.loc[:,'SAP编号'].astype('str')
	data = data[(data.loc[:,'SAP编号'].notnull())&(data.loc[:,'SAP编号'].str.isnumeric())]
	data = data.fillna(0)
	data.loc[:,'SAP编号'] = data.loc[:,'SAP编号'].astype('int')
	return data


#参数处理
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
		if "考勤统计" in j:
			kq = kq.append(pd.read_excel(files,sheet_name="考勤统计",header=1),ignore_index=True)
		if "津贴明细" in j:
			jt = jt.append(pd.read_excel(files,sheet_name="津贴明细"),ignore_index=True)
		if "社保统计" in j:
			ss = ss.append(pd.read_excel(files,sheet_name="社保统计",header=3),ignore_index=True)
		if "税务系统" in j:
			sw = sw.append(pd.read_excel(files,sheet_name="专项附加扣除-税务系统"),ignore_index=True)
		if "薪资数据集" in j:
			sap = sap.append(pd.read_excel(files,sheet_name="薪资数据集-Sap"),ignore_index=True)
		if "薪资异动表" in j:
			xz = xz.append(pd.read_excel(files,sheet_name="薪资异动表"),ignore_index=True)
		if "津贴异动表" in j:
			jt1 = jt1.append(pd.read_excel(files,sheet_name="津贴异动表"),ignore_index=True)
		if "小时工" in j:
			sl = sl.append(pd.read_excel(files,sheet_name="小时工"),ignore_index=True)
		if "银行" in j:
			bk = bk.append(pd.read_excel(files,sheet_name="银行明细",dtype={'银行代码':'str','银行账号':'str'}),ignore_index=True)





#考勤
kq = DataFrame(kq, columns=col_kq)
output(kq,"考勤")

#津贴明细
jt = DataFrame(jt, columns=col_jt)
output(jt,"津贴明细")

#社保
ss = DataFrame(ss, columns=col_ss)
output(ss,"社保")

#附加专项
stan(sw)
stan(sap)
if (len(sw.index) >= 1)&(len(sap.index) >= 1):
	fj = pd.merge(sw,sap,on='SAP编号',how='outer')
	fj.loc[:,'子女教育'] = fj.loc[:,'累计子女教育_x'] - fj.loc[:,'累计子女教育_y']
	fj.loc[:,'住房租金'] = fj.loc[:,'累计住房租金_x'] - fj.loc[:,'累计住房租金_y']
	fj.loc[:,'住房贷款'] = fj.loc[:,'累计住房贷款_x'] - fj.loc[:,'累计住房贷款_y']
	fj.loc[:,'赡养老人'] = fj.loc[:,'累计赡养老人_x'] - fj.loc[:,'累计赡养老人_y']
	fj.loc[:,'继续教育'] = fj.loc[:,'累计继续教育_x'] - fj.loc[:,'累计继续教育_y']
	fj = DataFrame(fj,columns=['SAP编号','子女教育','住房租金','住房贷款','赡养老人','继续教育'])
	output(fj,"附加专项")
else:
	print("未发现完整的附加专项相关数据!")

#薪资异动表
if len(xz) >= 1:
	xz = DataFrame(xz,columns=['SAP编号','姓名','当地最低工资标准','薪资'])
	output(xz,"薪资异动表")
else:
	print("未发现薪资异动数据!")

#津贴异动表
if len(jt1) >= 1:
	jt1 = DataFrame(jt1,columns=['SAP编号','姓名','项目','金额'])
	output(jt1,"津贴异动表")
else:
	print("未发现津贴异动数据!")

#小时工
if len(sl) >= 1:
	sl = DataFrame(sl,columns=['SAP编号','小时数','时薪','天数','日薪','提成','失货','劳务税'])
	output(stan(sl),"小时工")
else:
	print("未发现小时工数据!")

#银行明细
if len(bk) >= 1:
	bk = DataFrame(bk,columns=['SAP编号','银行代码','银行账号'])
	bk = bk.dropna(axis=0,how='any')
	bk.loc[:,'银行代码'] = bk.loc[:,'银行代码'].astype('str')
	bk.loc[:,'银行账号'] = bk.loc[:,'银行账号'].astype('str')
	output(bk,"银行明细")
else:
	print("未发现银行明细数据!")

print("\n中间表格已完成创建,请进行算薪操作,谢谢!")
input()


