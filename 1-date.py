#--coding:utf-8--

import pandas as pd
from pandas import DataFrame,Series
import numpy as np
import datetime
import os
import xlrd

print("欢迎使用合并表格模板\n一切解释权均归开发者所有!\n开发者: 凡凡\n")

#参数指定
path = 'C:\\Users\\hq01ug601\\Desktop\\工资明细'
path_mid = 'D:\\根目录\\项目\\中间表格'


#人事范围参数
dir_hrlim = {"人事范围":"范围","总部总直总监级以上":"SH","总部MB总监级以上":"SH","总部MC总监级以上":"SH",
			"总部邦购总监级以上":"SH","总部总直经理级类":"SH","总部MB经理级类":"SH","总部MC经理级类":"SH",
			"总部邦购经理级类":"SH","总部总直主管级类":"SH","总部MB主管级类":"SH","总部MC主管级类":"SH",
			"总部邦购主管级类":"SH","总部总直员工":"SH","总部MB员工":"SH","总部MC员工":"SH",
			"总部邦购员工":"SH","总部总直驻外":"SH","总部MB驻外":"SH","总部MC驻外":"SH",
			"总部邦购驻外":"SH","总部特殊人员":"SH","总部MB区域驻外":"SH","总部MC区域驻外":"SH",
			"总部邦购区域驻外":"SH","总部总直区域驻外":"SH","上海总直":"SH","上海MB":"SH",
			"上海MC":"SH","苏州总直":"SU","苏州MB":"SU","苏州MC":"SU","南京总直":"NJ",
			"南京MB":"NJ","南京MC":"NJ","合肥总直":"HF","合肥MB":"HF","合肥MC":"HF",
			"杭州总直":"HZ","杭州MB":"HZ","杭州MC":"HZ","宁波总直":"NB","宁波MB":"NB",
			"宁波MC":"NB","温州总直":"WZ","温州MB":"WZ","温州MC":"WZ","温州区域配送":"WZ",
			"北京总直":"BJ","北京MB":"BJ","北京MC":"BJ","天津总直":"TJ","天津MB":"TJ",
			"天津MC":"TJ","天津区域配送":"TJ","济南总直":"JN","济南MB":"JN","济南MC":"JN",
			"哈尔滨总直":"HE","哈尔滨MB":"HE","哈尔滨MC":"HE","长春总直":"CC","长春MB":"CC",
			"长春MC":"CC","沈阳总直":"SY","沈阳MB":"SY","沈阳MC":"SY","沈阳区域配送":"SY",
			"太原总直":"TY","太原MB":"TY","太原MC":"TY","石家庄总直":"SJ","石家庄MB":"SJ",
			"石家庄MC":"SJ","郑州总直":"ZZ","郑州MB":"ZZ","郑州MC":"ZZ","西安总直":"SX",
			"西安MB":"SX","西安MC":"SX","西安区域配送":"SX","兰州总直":"LZ","兰州MB":"LZ",
			"兰州MC":"LZ","乌鲁木齐总直":"WQ","乌鲁木齐MB":"WQ","乌鲁木齐MC":"WQ","成都总直":"CD",
			"成都MB":"CD","成都MC":"CD","成都区域配送":"CD","重庆总直":"CQ","重庆MB":"CQ",
			"重庆MC":"CQ","昆明总直":"KM","昆明MB":"KM","昆明MC":"KM","广州区域配送":"GZ",
			"广州总直":"GZ","广州MB":"GZ","广州MC":"GZ","深圳总直":"SZ","深圳MB":"SZ",
			"深圳MC":"SZ","南宁总直":"NN","南宁MB":"NN","南宁MC":"NN","武汉总直":"WH",
			"武汉MB":"WH","武汉MC":"WH","武汉区域配送":"WH","南昌总直":"NC","南昌MB":"NC",
			"南昌MC":"NC","福州总直":"FZ","福州MB":"FZ","福州MC":"FZ","东莞总直":"DG",
			"东莞MB":"DG","东莞MC":"DG","长沙总直":"CS","长沙MB":"CS","长沙MC":"CS",
			"贵阳总直":"GY","贵阳MB":"GY","贵阳MC":"GY","青岛总直":"QD","青岛MB":"QD",
			"青岛MC":"QD","内蒙古总直":"NM","内蒙古MB":"NM"}

#经常性参数


root = 'D:\\根目录\\人事异动\\'
month = (datetime.datetime.now() - datetime.timedelta(30,0,0,0)).month
file800 = root + str(month) + "月人事异动表-800.xlsx"
file830 = root + str(month) + "月人事异动表-830.xlsx"

if (os.path.exists(file800)) & (os.path.exists(file800)):
	df = pd.read_excel(file800).append(pd.read_excel(file830))
elif os.path.exists(file800):
	df = pd.read_excel(file800)
elif os.path.exists(file830):
	df = pd.read_excel(file830)
else:
	df = DataFrame()
	print("未找到当月相关人事异动表!")
	
df = df.reset_index()

def stan(data):
	for i in range(len(data.index)):
		data.loc[i,'月初日期'] = datetime.datetime(int(data.loc[i,'年']),int(data.loc[i,'月']),1)
		data.loc[i,'月末日期'] = datetime.datetime(int(data.loc[i,'年']),int(data.loc[i,'月'])+1,1) - datetime.timedelta(1,0,0,0)
		data.loc[i,'初始日期'] = data.loc[i,'月末日期'].strftime("%Y%m%d")
		if data.loc[i,'入职'] == '新员工入职':
			data.loc[i,'开始日期'] = data.loc[i,'入职日期'].strftime("%Y%m%d")
		else:
			data.loc[i,'开始日期'] = data.loc[i,'月初日期'].strftime("%Y%m%d")
	data.loc[:,'结束日期'] = "9991231"
	return data

if len(df) >= 1:
	hrlim = DataFrame(Series(dir_hrlim),columns=['范围'])
	stan(df)
	data = pd.merge(df, hrlim, left_on='人事范围描述', right_index=True, how='left')
	data.loc[:,'系统'] = 800
	data.loc[data['SAP人员编号'] > 6000000,"系统"] = 830
	data_end = data.loc[:,["SAP人员编号","开始日期","初始日期","结束日期","范围","系统"]]
	file_name = 'D:\\根目录\\项目\\中间表格\\日期参数\\' + str(month) +'月开始日期整理表.xlsx'
	data_end.to_excel(file_name, index=False)
	print("开始日期整理表已成功导出!")
else:
	print("开始日期整理表未导出,请核对人事异动表信息!")

input()













