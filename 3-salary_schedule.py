#--coding:utf-8--

import pandas as pd
from pandas import DataFrame,Series
import numpy as np
import datetime
import os
import xlrd

print("欢迎使用合并表格模板\n一切解释权均归开发者所有!\n开发者: 凡凡\n")

#参数指定
desk = os.path.join(os.path.expanduser("~"),"Desktop")

path = 'D:\\根目录\\项目\\中间表格\\日期参数\\'
path_mid = 'D:\\根目录\\项目\\中间表格\\'
path_text = desk + "\\1-text\\"
path_char = desk + "\\2-split\\"

month = (datetime.datetime.now()).month-1
path_date = path + str(month) + '月开始日期整理表.xlsx'

df_date = pd.read_excel(path_date)


#工资项参数
dir_or = {"附加医疗缴纳金额个人":"0917","附加医疗缴纳金额公司":"0918","养老补缴金额个人":"0919",
		  "养老补缴金额公司":"0920","医疗补缴金额个人":"0921","医疗补缴金额公司":"0922",
		  "失业补缴金额个人":"0923","失业补缴金额公司":"0924","工伤补缴金额公司":"0926",
		  "生育补缴金额公司":"0928","公积金补缴金额个人":"0929","公积金补缴金额公司":"0930",
		  "附加医疗补缴金额个人":"0931","附加医疗补缴金额公司":"0932","综合保险公司补缴金额":"0937",
		  "综合保险个人补缴金额":"0939","小时工工资":"1004","饭贴":"1202","其他津贴":"1211",
		  "独生子女费":"1213","值班津贴":"1216","津贴1":"1217","津贴2":"1218","全勤奖":"1304",
		  "销售之星":"1306","服务之星":"1307","长期服务奖":"1311","其他奖金":"1314","业绩提成实得值":"1316",
		  "客单件奖金":"1317","月度考核奖金(计年终奖)":"1320","月度考核奖金(不计年终奖)":"1321",
		  "婚育津贴":"1322","业绩奖金（不计年终奖）":"1323","业绩奖金（计年终奖）":"1324","月度个人评优奖金":"1325",
		  "非月度奖金":"1326","奖金1":"1327","奖金2":"1328","调薪补款":"2101","考勤补款":"2102",
		  "产假补款":"2103","税前其他补款":"2108","住房补贴":"2111","社保个人补款":"2112","公积金个人补款":"2113",
		  "补偿金":"2201","税后其他补款":"2203","补扣社保":"2204","补扣公积金":"2205","税前其他扣款":"2301",
		  "考勤扣款":"2302","病假扣款":"2303","住宿扣款":"2401","财务扣款":"2402","失货扣款":"2403",
		  "工会会费":"2404","税后其他扣款":"2405","住宿水电费及饮用水扣款":"2410","住宿费扣款-总部六灶住宿及水电费":"2411",
		  "住宿费扣款-总部六灶饮用水费":"2412","住宿费扣款(梅花苑)":"2413","子女教育":"/4J1",
		  "住房贷款":"/4J4","住房租金":"/4J5","赡养老人":"/4J6","继续教育":"/4J2","业绩奖金(不计年终奖)":"1323",
		  "业绩奖金(计年终奖)":"1324","提成":"1324","失货":"2403","劳务税":"2405"}
df_or = DataFrame(Series(dir_or),columns=['工资项'])

dir_jc = {"养老缴纳金额个人":"901","养老缴纳金额公司":"902","医疗缴纳金额个人":"903",
		  "医疗缴纳金额公司":"904","失业缴纳金额个人":"905","失业缴纳金额公司":"906",
		  "工伤缴纳金额公司":"907","生育缴纳金额公司":"908","公积金缴纳金额个人":"909",
		  "综合保险公司缴纳金额":"912","营业税（南京）":"940","综合保险个人缴纳金额":"952",
		  "公积金缴纳金额公司":"973","当地最低工资标准":"1006","驻外津贴":"1201",
		  "饭贴":"1202","大店津贴":"1212","职级津贴":"1214","岗位津贴":"1215",
		  "年终奖基数":"1312","服装费":"1318","奖金标准":"1329","工会会费":"2404",
		  "政策性个人免税标准":"2800","税基调整项":"2901"}
df_jc = DataFrame(Series(dir_jc),columns=['工资项'])

dir_kq = {"当月计薪天数":"3102","应出勤天数":"3101","实际出勤天数":"3104","病假天数":"3203",
		  "无薪事假天数":"3202","产假天数":"3204","婚假天数":"3206","丧假天数":"3207",
		  "空勤天数":"3201","旷工天数":"3209","出差天数":"3210","陪产假天数":"3205",
		  "工伤假天数":"3208","店铺未打卡（次）":"3300","非店铺未打卡（次）":"3301",
		  "非店铺迟到0-30M":"3302","非店铺迟到31-60M":"3303","非店铺迟到61-120M":"3304",
		  "非店铺迟到120M以上":"3305","行政早退30M以内":"3317","行政早退30M以上":"3306",
		  "店铺迟到0-10M":"3307","店铺迟到11-30M":"3308","店铺迟到31-60M":"3309",
		  "店铺迟到61-120M":"3310","店铺迟到120M以上":"3311","店铺早退1小时内":"3312",
		  "店铺早退1小时以上":"3313","平时加班时":"3401","节日加班时":"3403","周末加班时":"3402"}
df_kq = DataFrame(Series(dir_kq),columns=['工资项'])


#自定义函数
def opt(df, name, form='.xls'):
	df_800 = df.loc[df.loc[:,'系统'] == 800,:]
	df_830 = df.loc[df.loc[:,'系统'] == 830,:]
	del df_800['系统']
	del df_830['系统']
	if len(df_800) >= 1:
		df_800.sort_values(by=['SAP编号'])
	if len(df_830) >= 1:
		df_830.sort_values(by=['SAP编号'])
	if form == '.xls':
		if len(df_800) >= 1:
			df_800.to_excel(path_char + name + "800" + form,index=False)
			print(name + "800已导出,条数为: " + str(len(df_800)))
		if len(df_830) >= 1:
			df_830.to_excel(path_char + name + "830" + form,index=False)
			print(name + "830已导出,条数为: " + str(len(df_830)))
	elif form == '.txt':
		if len(df_800) >= 1:
			df_800.loc[:,'开始日期'] = df_800.loc[:,'开始日期'].astype('int')
			df_800.to_csv(path_text + name + "800" + form,index=False,sep='\t')
			print(name + "800已导出,条数为: " + str(len(df_800)))
		if len(df_830) >= 1:
			df_830.loc[:,'开始日期'] = df_830.loc[:,'开始日期'].astype('int')
			df_830.to_csv(path_text + name + "830" + form,index=False,sep='\t')
			print(name + "830已导出,条数为: " + str(len(df_830)))
	else:
		print("未按指定格式提交参数!")


#考勤
if os.path.exists(path_mid + '考勤.xlsx'):
	df_atd = pd.read_excel(path_mid + '考勤.xlsx')
	df_atd = df_atd.melt(id_vars=['SAP编号','姓名'], var_name="属性", value_name="时数")
	df_atd = pd.merge(df_atd, df_kq, left_on='属性', right_index=True, how='left')
	df_atd = pd.merge(df_atd, df_date, left_on='SAP编号', right_on='SAP人员编号', how='left')
	df_atd.loc[:,'时数'] = df_atd.loc[:,'时数'].astype('float')
	df_attendance = DataFrame(df_atd[df_atd['时数'] > 0],columns=['SAP编号','姓名','工资项','初始日期','时数','编号','单位','金额','系统'])
	
	opt(df_attendance,"考勤")
else:
	print("未发现考勤数据!")


#津贴明细
if os.path.exists(path_mid + '津贴明细.xlsx'):
	df_jt = pd.read_excel(path_mid + '津贴明细.xlsx')
	df_jt.rename(columns={'工资项':'属性'},inplace=True)
	df_jt.loc[:,'金额'] = df_jt.loc[:,'金额'].apply(lambda x:round(x,2))
	df_jt = pd.pivot_table(df_jt,index=['SAP编号','姓名','属性'], values=['金额'],aggfunc='sum').reset_index()
	df_jt.loc[df_jt['属性'].str.contains("扣"),"金额"] = df_jt.loc[df_jt['属性'].str.contains("扣"),"金额"].apply(lambda x:-abs(x))
else:
	df_jt = DataFrame()


#社保
if os.path.exists(path_mid + '社保.xlsx'):
	df_ss = pd.read_excel(path_mid + '社保.xlsx')
	df_ss = pd.merge(df_ss,df_date,left_on='SAP编号',right_on='SAP人员编号',how='left')
	
	df_kg = DataFrame(df_ss[(df_ss['社保账户']!=0)|(df_ss['公积金账户']!=0)],columns=['SAP编号','姓名'])
	df_kg['0001'] = 'ZM'
	df_kg['0002'] = 'ZM'
	df_kg['0003'] = 'ZM'
	df_kg['0004'] = 'ZM'
	df_kg['0005'] = 'ZM'
	df_kg = df_kg.melt(id_vars=['SAP编号','姓名'],var_name="子信息类型",value_name='分摊范围')
	df_kg = pd.merge(df_kg, df_ss, on=['SAP编号','姓名'],how='left')
	
	df_kg.loc[:,'分摊标准'] = "01"
	df_kg.loc[(df_kg.loc[:,'子信息类型']=='0001')&(df_kg.loc[:,'养老缴纳金额个人']==0)&(df_kg.loc[:,'养老缴纳金额公司']!=0),"分摊标准"] = '02'
	df_kg.loc[(df_kg.loc[:,'子信息类型']=='0002')&(df_kg.loc[:,'失业缴纳金额个人']==0)&(df_kg.loc[:,'失业缴纳金额公司']!=0),"分摊标准"] = '02'
	df_kg.loc[(df_kg.loc[:,'子信息类型']=='0003')&(df_kg.loc[:,'医疗缴纳金额个人']==0)&(df_kg.loc[:,'医疗缴纳金额公司']!=0),"分摊标准"] = '02'
	df_kg.loc[(df_kg.loc[:,'子信息类型']=='0003')&(df_kg.loc[:,'医疗缴纳金额个人']==0)&(df_kg.loc[:,'医疗缴纳金额公司']!=0),"分摊标准"] = '02'

	df_kg.loc[:,'雇主和雇员支付的分摊'] = ""
	df_kg.loc[(df_kg.loc[:,'子信息类型']=='0001')&(df_kg.loc[:,'养老缴纳金额个人'] + df_kg.loc[:,'养老缴纳金额公司'] > 0),"雇主和雇员支付的分摊"] = "X"
	df_kg.loc[(df_kg.loc[:,'子信息类型']=='0002')&(df_kg.loc[:,'失业缴纳金额个人'] + df_kg.loc[:,'失业缴纳金额公司'] > 0),"雇主和雇员支付的分摊"] = "X"
	df_kg.loc[(df_kg.loc[:,'子信息类型']=='0003')&(df_kg.loc[:,'医疗缴纳金额个人'] + df_kg.loc[:,'医疗缴纳金额公司'] > 0),"雇主和雇员支付的分摊"] = "X"
	df_kg.loc[(df_kg.loc[:,'子信息类型']=='0004')&(df_kg.loc[:,'工伤缴纳金额公司'] > 0),"雇主和雇员支付的分摊"] = "X"
	df_kg.loc[(df_kg.loc[:,'子信息类型']=='0005')&(df_kg.loc[:,'生育缴纳金额公司'] > 0),"雇主和雇员支付的分摊"] = "X"
	
	df_kg.loc[:,'无分摊'] = "X"
	df_kg.loc[df_kg.loc[:,'雇主和雇员支付的分摊'] == "X", "无分摊"] = ""
	df_kg.loc[:,'分摊更改原因代码'] = '03'
	df_kg.loc[df_kg.loc[:,'无分摊'] == "X", '分摊更改原因代码'] = '14'
	df_kg.loc[:,'行业'] = '01'
	df_kg.loc[:,'分摊组'] = "ZM01"
	
	df_kg_end = DataFrame(df_kg,columns=['SAP编号', '子信息类型','开始日期','分摊更改原因代码','保险单编号','分摊范围','行业','分摊组','分摊标准','雇主和雇员支付的分摊','仅雇主的分摊支付','无分摊','系统'])
	opt(df_kg_end, "社保分摊", '.txt')
	
	
	df_kg_gjj = DataFrame(df_ss[df_ss['公积金账户'] != 0])
	df_kg_gjj.loc[:,'地区'] = "ZM"
	df_kg_gjj.loc[:,'组'] = "ZM01"
	df_kg_gjj.loc[:,'级别'] = "01"

	df_kg_gjj.loc[:,'雇主和雇员缴纳'] = ""
	df_kg_gjj.loc[df_kg_gjj.loc[:,'公积金缴纳金额个人'] + df_kg_gjj.loc[:,'公积金缴纳金额公司'] > 0,"雇主和雇员缴纳"] = "X"
	df_kg_gjj.loc[:,'不缴纳'] = "X"
	df_kg_gjj.loc[df_kg_gjj.loc[:,'雇主和雇员缴纳'] == "X","不缴纳"] = ""
	df_kg_gjj_end = DataFrame(df_kg_gjj, columns=['SAP编号','开始日期','住房公积金账号','地区','组','级别','雇主和雇员缴纳','单位缴纳','员工缴纳','不缴纳','系统'])
	opt(df_kg_gjj_end, "公积金", '.txt')
	
	
	df_ss_mid = pd.read_excel(path_mid + '社保.xlsx')
	del df_ss_mid['社保账户']
	del df_ss_mid['公积金账户']
	df_ss_mid = df_ss_mid.melt(id_vars=['SAP编号','姓名'],var_name='属性',value_name='金额')
else:
	df_ss_mid = DataFrame()


#附加专项
if os.path.exists(path_mid + '附加专项.xlsx'):
	df_fj = DataFrame(pd.read_excel(path_mid + '附加专项.xlsx'),columns=['SAP编号','姓名','子女教育','住房租金','住房贷款','赡养老人','继续教育'])
	df_fj = df_fj.melt(id_vars=['SAP编号','姓名'], var_name="属性", value_name="金额")
else:
	df_fj = DataFrame()


#薪资异动表
if os.path.exists(path_mid + '薪资异动表.xlsx'):
	df_xz = pd.read_excel(path_mid + '薪资异动表.xlsx')
	df_xz.loc[:,'更改原因'] = '50'
	df_xz.loc[:,'工资等级类型'] = "01"
	df_xz.loc[:,'级别'] = '0001'
	df_xz.loc[:,'档次'] = 'A0'
	df_xz.loc[:,'基本工资工资项'] = 1001
	df_xz.loc[:,'基本工资金额'] = df_xz.loc[:,'当地最低工资标准']
	df_xz.loc[:,'职级工资工资项'] = 1002
	df_xz.loc[:,'职级工资金额'] = df_xz.loc[:,'薪资'] - df_xz.loc[:,'当地最低工资标准']
	df_xz.loc[:,'固定工资标准工资项'] = 1003
	df_xz.loc[:,'固定工资'] = df_xz.loc[:,'薪资']
	df_xz = pd.merge(df_xz, df_date,left_on='SAP编号',right_on='SAP人员编号',how='left')
	df_xz_end = DataFrame(df_xz[df_xz['薪资']>0],columns = ['SAP编号','姓名','开始日期','结束日期','更改原因','工资等级类型','范围','级别','档次','基本工资工资项',\
						  '基本工资金额','职级工资工资项','职级工资金额','固定工资标准工资项','固定工资','小时工工资项目','系统'])
	opt(df_xz_end,"薪资",'.xls')
	
	
	df_xz_dj = DataFrame(df_xz[df_xz['薪资']==0],columns=['SAP人员编号','开始日期','系统'])
	opt(df_xz_dj,"薪资定界",'.txt')
	
	
	df_xz_jc = DataFrame(df_xz, columns=['SAP编号','姓名','当地最低工资标准'])
	df_xz_jc = df_xz_jc.melt(id_vars=['SAP编号','姓名'], var_name='属性', value_name='金额')
else:
	df_xz_jc = DataFrame()


#津贴异动表
if os.path.exists(path_mid + '津贴异动表.xlsx'):
	df_jty = pd.read_excel(path_mid + '津贴异动表.xlsx')
	df_jty.rename(columns={'项目':'属性'}, inplace=True)
else:
	df_jty = DataFrame()


#小时工
if os.path.exists(path_mid + '小时工.xlsx'):
	df_sl = DataFrame(pd.read_excel(path_mid + '小时工.xlsx'),columns=['SAP编号','姓名','小时数','时薪','天数','日薪','提成','失货','劳务税'])
	
	df_sl.fillna(0,inplace=True)
	df_sl.loc[:,'小时工工资'] = df_sl.loc[:,'小时数'] * df_sl.loc[:,'时薪'] + df_sl.loc[:,'天数'] * df_sl.loc[:,'日薪']
	df_sl = DataFrame(df_sl, columns=['SAP编号','姓名','小时工工资','提成','失货','劳务税'])
	df_sl = pd.pivot_table(df_sl, index=['SAP编号'], values=['小时工工资','提成','失货','劳务税'], aggfunc='sum').reset_index()

	df_sl_end = df_sl.melt(id_vars=['SAP编号'], var_name="属性", value_name='金额')
	df_sl_end = DataFrame(df_sl_end, columns=['SAP编号','姓名','属性','金额'])
else:
	df_sl_end = DataFrame()


#银行明细
if os.path.exists(path_mid + '银行明细.xlsx'):
	df_bk = DataFrame(pd.read_excel(path_mid + '银行明细.xlsx'))
	df_bk.loc[:,'主要银行'] = str(0)
	df_bk = pd.merge(df_bk, df_date, left_on='SAP编号', right_on='SAP人员编号', how='left')
	df_bk_end = DataFrame(df_bk, columns=['SAP编号','开始日期','结束日期','主要银行','姓名','银行代码','银行账号','系统'])
	opt(df_bk_end, "银行明细",'.txt')
else:
	print("未发现银行数据信息!")


#经常性整合
df_jcx = pd.concat([df_ss_mid,df_jty,df_xz_jc], axis=0)
df_jcx = pd.merge(df_jcx, df_date, left_on='SAP编号',right_on='SAP人员编号',how='left')
df_jcx = pd.merge(df_jcx, df_jc, left_on='属性', right_index=True,how='left')
df_jcx = df_jcx[df_jcx['工资项'].notnull()]
df_jcx.loc[:,"工资项"] = df_jcx.loc[:,"工资项"].astype('int').astype('str')
df_jcx.loc[:,"工资项"] = df_jcx.loc[:,"工资项"].apply(lambda x:x.zfill(4))

df_jcx_opt = DataFrame(df_jcx[(df_jcx['金额']>0)&(df_jcx['工资项'].notnull())],columns=['SAP编号','姓名','开始日期','结束日期','工资项','金额','编号','单位','系统'])
opt(df_jcx_opt, "经常性", '.xls')

df_jcx_dj = DataFrame(df_jcx[(df_jcx['金额']==0)&(df_jcx['工资项'].notnull())],columns=['SAP编号','工资项','开始日期','系统'])
df_jcx_dj = DataFrame(df_jcx_dj[(df_jcx_dj['工资项']=="1214")|(df_jcx_dj['工资项']=="1202")|(df_jcx_dj['工资项']=="1201")|(df_jcx_dj['工资项']=="1212")|(df_jcx_dj['工资项']=="1215")])
opt(df_jcx_dj, "经常性定界", '.txt')


#偶然性整合
df_orx = pd.concat([df_jt,df_ss_mid,df_fj,df_sl_end],axis=0)
df_orx = pd.merge(df_orx, df_date, left_on='SAP编号',right_on='SAP人员编号',how='left')
df_orx = pd.merge(df_orx, df_or, left_on='属性', right_index=True,how='left')
df_orx = df_orx[df_orx['工资项'].notnull()]
df_orx.loc[:,'工资项'] = df_orx.loc[:,'工资项'].astype('str')
df_orx.loc[df_orx['工资项'].str.find(".")!=-1,"工资项"] = df_orx.loc[df_orx['工资项'].str.find(".")!=-1,"工资项"].astype('int').astype('str')
df_orx.loc[:,"工资项"] = df_orx.loc[:,"工资项"].apply(lambda x:x.zfill(4))

df_orx_opt = DataFrame(df_orx[(df_orx['金额']>0)&(df_orx['工资项'].notnull())&(df_orx['初始日期'].notnull())],columns=['SAP编号','姓名','工资项','金额','编号','单位','初始日期','分配编号','系统'])
df_orx_opt.loc[:,'初始日期'] = df_orx_opt.loc[:,'初始日期'].astype('int')
df_orx_opt.loc[:,'系统'] = df_orx_opt.loc[:,'系统'].astype('int')
opt(df_orx_opt,"偶然性", '.xls')

print("\n模板表已经制作完成,请导入系统!")
input()
