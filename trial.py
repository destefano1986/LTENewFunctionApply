# -*- coding:utf-8 -*-
import pandas as pd
import numpy as np
import xlrd

def read_xls(xlsname, sheetname):
    fname = xlsname
    bk = xlrd.open_workbook(fname)
    shxrange = range(bk.nsheets)
    #读取excel文件，并将地址进行保存
    try:
        sh = bk.sheet_by_name(sheetname)
    except:
        print ("no sheet in %s named Sheet1" % fname)
    print ("finishing "+xlsname)
    return sh

paras = read_xls("20180423.xlsx", "20180423")
parasnrows = paras.nrows
parasncols = paras.ncols
#获取各列数据第2至18列
row_list = []
for i in range(1,parasnrows):
    row_data = paras.row_values(i)
    row_list.append(row_data)
df1 = pd.DataFrame(row_list,columns=['时间','CGI','区县','区域维护部','厂家名称','所属E-NODEB','小区中文名',\
                                    '小区英文名','经度','纬度','设备维护状态','管理状态','覆盖类型','覆盖场景',\
                                    '上下行子帧配置','特殊子帧模式','小区带宽','工作频段','中心载频的信道号',\
                                    '载频数量','跟踪区码','物理小区识别码','本地小区标识','电子下倾角','机械下倾角',\
                                    '总下倾角','方位角','天线挂高','PRACH配置索引','参考信号功率','LTE小区的LTE邻区数量',\
                                    'LTE小区的TD邻区数量','LTE小区的GSM邻区数量','所属网格编号','地址','stationcatalog-基站属性',\
                                    '交维状态','自动路测区域','属地分公司','eNodeBID'])
for i in range(1, parasnrows - 1):
    print('breakpoint: xCGI ' + str(i))
    df1.loc[i, 'xCGI'] = '460-00-' + str(df1.loc[i, 'eNodeBID']) + '-' + str(df1.loc[i, '本地小区标识'])
df1 = df1[['时间','CGI','xCGI','区县','区域维护部','厂家名称','所属E-NODEB','小区中文名',\
                                    '小区英文名','经度','纬度','设备维护状态','管理状态','覆盖类型','覆盖场景',\
                                    '上下行子帧配置','特殊子帧模式','小区带宽','工作频段','中心载频的信道号']]
print (df1)