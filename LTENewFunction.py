# -*- coding:utf-8 -*-
import pandas as pd
import numpy as np
import xlrd
import datetime

#读取excel文件和excel工参文件
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


#制作各项新功能清洗(DLCOMP、ULCOMP、uplink_data_compression、dl_qam、IR)
def ul_dl_comp():
    compnrows = cellalgoswitch.nrows
    compncols = cellalgoswitch.ncols
    comp_list = []
    print ('breakpoint one: comp')
    for i in range(1,compnrows):
        print('breakpoint two: comp '+str(i))
        comp_data = cellalgoswitch.row_values(i)
        comp_list.append(comp_data)
    comp = pd.DataFrame(comp_list)
    print (comp.columns)
    comp = comp[[1,2,38,41,86,92,83,52,22]]
    comp.columns = ['SITEID','LocalCellId','DlCompSwitch','UplinkCompSwitch','Dl256QamAlgoSwitch','UdcAlgoSwitch','UlJRAntNumCombSw','InterfRandSwitch','UlSchSwitch']
    comp.drop(comp.index[0],inplace=True)
    comp['SITEID'] = comp['SITEID'].astype('int')
    comp['LocalCellId'] = comp['LocalCellId'].astype('int')
    for i in range(1,compnrows-1):
        print('breakpoint three: comp ' + str(i))
        comp.loc[i,'CGI']='460-00-'+str(comp.loc[i,'SITEID'])+'-'+str(comp.loc[i,'LocalCellId'])
        if 'IntraDlCompSwitch:On' in comp.loc[i,'DlCompSwitch']:
            comp.loc[i,'DlComp'] = 'Y'
        else:
            comp.loc[i,'DlComp'] = 'N'
        if 'UlJointReceptionSwitch:On;UlJointReceptionPhaseIISwitch:On;' in comp.loc[i,'UplinkCompSwitch']:
            comp.loc[i,'UlComp'] = 'Y'
        else:
            comp.loc[i,'UlComp'] = 'N'
        if 'Dl256QamSwitch:On' in comp.loc[i,'Dl256QamAlgoSwitch']:
            comp.loc[i,'DL 256QAM'] = 'Y'
        else:
            comp.loc[i,'DL 256QAM'] = 'N'
        if 'BasicUdcSwitch:On' in comp.loc[i,'UdcAlgoSwitch']:
            comp.loc[i,'UDC'] = 'Y'
        else:
            comp.loc[i,'UDC'] = 'N'
        if 'ENB_BASED' in comp.loc[i,'InterfRandSwitch']:
            comp.loc[i,'IF'] = 'Y'
        else:
            comp.loc[i,'IF'] = 'N'
        if 'UlVmimoSwitch:On' in comp.loc[i,'UlSchSwitch']:
            comp.loc[i,'VMIMO'] = 'Y'
        else:
            comp.loc[i,'VMIMO'] = 'N'
    print('breakpoint four: comp')
    comp.drop('SITEID',axis=1,inplace=True)
    comp.drop('LocalCellId', axis = 1, inplace = True)
    comp.drop('DlCompSwitch', axis = 1, inplace = True)
    comp.drop('UplinkCompSwitch', axis = 1, inplace = True)
    comp.drop('Dl256QamAlgoSwitch', axis=1, inplace=True)
    comp.drop('UdcAlgoSwitch', axis=1, inplace=True)
    comp.drop('UlJRAntNumCombSw', axis=1, inplace=True)
    comp.drop('InterfRandSwitch', axis=1, inplace=True)
    comp.drop('UlSchSwitch', axis=1, inplace=True)
    print('breakpoint five: comp')
    return comp

#制作各项新功能清洗(ULCa)
def ul_ca():
    ulcanrows = camctcfg.nrows
    ulcancols = camctcfg.ncols
    ulca_list = []
    print('breakpoint one: ul_ca')
    for i in range(1,ulcanrows):
        print('breakpoint two: ul_ca ' + str(i))
        ulca_data = camctcfg.row_values(i)
        ulca_list.append(ulca_data)
    ulca = pd.DataFrame(ulca_list)
    ulca = ulca[[1,2,16]]
    ulca.columns = ['SITEID','LocalCellId','CellCaAlgoSwitch']
    ulca.drop(ulca.index[0],inplace=True)
    ulca['SITEID'] = ulca['SITEID'].astype('int')
    ulca['LocalCellId'] = ulca['LocalCellId'].astype('int')
    for i in range(1,ulcanrows-1):
        print('breakpoint three: ul_ca ' + str(i))
        ulca.loc[i,'CGI']='460-00-'+str(ulca.loc[i,'SITEID'])+'-'+str(ulca.loc[i,'LocalCellId'])
        if 'CaUl2CCSwitch:On' in ulca.loc[i,'CellCaAlgoSwitch']:
            ulca.loc[i,'UlCa'] = 'Y'
        else:
            ulca.loc[i,'UlCa'] = 'N'
    ulca.drop('SITEID',axis=1,inplace=True)
    ulca.drop('LocalCellId', axis = 1, inplace = True)
    ulca.drop('CellCaAlgoSwitch', axis = 1, inplace = True)
    return ulca

#制作各项新功能清洗(dLCa)
def dl_ca():
    dlcanrows = celldlschalgo.nrows
    dlcancols = celldlschalgo.ncols
    dlca_list = []
    print('breakpoint one: dl_ca')
    for i in range(1,dlcanrows):
        print('breakpoint two: dl_ca ' + str(i))
        dlca_data = celldlschalgo.row_values(i)
        dlca_list.append(dlca_data)
    dlca = pd.DataFrame(dlca_list)
    dlca = dlca[[1,2,6]]
    dlca.columns = ['SITEID','LocalCellId','CaSchStrategy']
    dlca.drop(dlca.index[0],inplace=True)
    dlca['SITEID'] = dlca['SITEID'].astype('int')
    dlca['LocalCellId'] = dlca['LocalCellId'].astype('int')
    for i in range(1,dlcanrows-1):
        print('breakpoint three: dl_ca ' + str(i))
        dlca.loc[i,'CGI']='460-00-'+str(dlca.loc[i,'SITEID'])+'-'+str(dlca.loc[i,'LocalCellId'])
        if 'differentiation schedule' in dlca.loc[i,'CaSchStrategy']:
            dlca.loc[i,'DlCa'] = 'Y'
        else:
            dlca.loc[i,'DlCa'] = 'N'
    dlca.drop('SITEID',axis=1,inplace=True)
    dlca.drop('LocalCellId', axis = 1, inplace = True)
    dlca.drop('CaSchStrategy', axis = 1, inplace = True)
    return dlca


#制作各项新功能清洗(high_speed)
def high_speed():
    hsnrows = cell.nrows
    hsncols = cell.ncols
    hs_list = []
    print('breakpoint one: high_speed')
    for i in range(1,hsnrows):
        print('breakpoint two: high_speed ' + str(i))
        hs_data = cell.row_values(i)
        hs_list.append(hs_data)
    hs = pd.DataFrame(hs_list)
    hs = hs[[1,2,23]]
    hs.columns = ['SITEID','LocalCellId','HighSpeedFlag']
    hs.drop(hs.index[0],inplace=True)
    hs['SITEID'] = hs['SITEID'].astype('int')
    hs['LocalCellId'] = hs['LocalCellId'].astype('int')
    for i in range(1,hsnrows-1):
        print('breakpoint three: high_speed ' + str(i))
        hs.loc[i,'CGI']='460-00-'+str(hs.loc[i,'SITEID'])+'-'+str(hs.loc[i,'LocalCellId'])
        if 'High speed cell flag' in hs.loc[i,'HighSpeedFlag']:
            hs.loc[i,'HighSpeed'] = 'Y'
        else:
            hs.loc[i,'HighSpeed'] = 'N'
    hs.drop('SITEID',axis=1,inplace=True)
    hs.drop('LocalCellId', axis = 1, inplace = True)
    hs.drop('HighSpeedFlag', axis = 1, inplace = True)
    return hs


#制作各项新功能清洗(64QAM)
def ul_qam():
    ulqamnrows = puschcfg.nrows
    ulqamncols = puschcfg.ncols
    ulqam_list = []
    print('breakpoint one: ul_qam')
    for i in range(1,ulqamnrows):
        print('breakpoint two: ul_qam ' + str(i))
        ulqam_data = puschcfg.row_values(i)
        ulqam_list.append(ulqam_data)
    ulqam = pd.DataFrame(ulqam_list)
    ulqam = ulqam[[1,2,10]]
    ulqam.columns = ['SITEID','LocalCellId','Qam64Enabled']
    ulqam.drop(ulqam.index[0],inplace=True)
    ulqam['SITEID'] = ulqam['SITEID'].astype('int')
    ulqam['LocalCellId'] = ulqam['LocalCellId'].astype('int')
    for i in range(1,ulqamnrows-1):
        print('breakpoint three: ul_qam ' + str(i))
        ulqam.loc[i,'CGI']='460-00-'+str(ulqam.loc[i,'SITEID'])+'-'+str(ulqam.loc[i,'LocalCellId'])
        if ulqam.loc[i,'Qam64Enabled'] == 'True':
            ulqam.loc[i,'UL 64QAM'] = 'Y'
        else:
            ulqam.loc[i,'UL 64QAM'] = 'N'
    ulqam.drop('SITEID',axis=1,inplace=True)
    ulqam.drop('LocalCellId', axis = 1, inplace = True)
    ulqam.drop('Qam64Enabled', axis = 1, inplace = True)
    return ulqam



if __name__ == "__main__":
    # 起始执行时间
    start = datetime.datetime.now()

    # 读取文件参数内容
    paras = read_xls("20180618.xlsx", "20180618")
    camctcfg = read_xls("CAMGTCFG.xlsx", "CAMGTCFG")
    cell = read_xls("CELL.xlsx", "CELL")
    cellalgoswitch = read_xls("CELLALGOSWITCH.xlsx", "CELLALGOSWITCH")
    celldlschalgo = read_xls("CELLDLSCHALGO.xlsx", "CELLDLSCHALGO")
    # cellulschalgo = read_xls("CELLULSCHALGO.xlsx", "CELLULSCHALGO")
    puschcfg = read_xls("PUSCHCFG.xlsx", "PUSCHCFG")

    #工参表内容输出
    parasnrows = paras.nrows
    parasncols = paras.ncols
    #获取各列数据第2至18列
    row_list = []
    for i in range(1,parasnrows):
        row_data = paras.row_values(i)
        row_list.append(row_data)
    df1 = pd.DataFrame(row_list,columns=['时间','CGI','eCGI','区县','区域维护部','厂家名称','所属E-NODEB','小区中文名',\
                                        '小区英文名','经度','纬度','设备维护状态','管理状态','覆盖类型','覆盖场景',\
                                        '上下行子帧配置','特殊子帧模式','小区带宽','工作频段','中心载频的信道号',\
                                        '载频数量','跟踪区码','物理小区识别码','本地小区标识','电子下倾角','机械下倾角',\
                                        '总下倾角','方位角','天线挂高','PRACH配置索引','参考信号功率','LTE小区的LTE邻区数量',\
                                        'LTE小区的TD邻区数量','LTE小区的GSM邻区数量','所属网格编号','地址','stationcatalog-基站属性',\
                                        '交维状态','自动路测区域','属地分公司','eNodeBID'])
    #for i in range(1, parasnrows - 1):
        #print('breakpoint: xCGI ' + str(i))
        #df1.loc[i, 'xCGI'] = '460-00-' + str(df1.loc[i, 'eNodeBID']) + '-' + str(df1.loc[i, '本地小区标识'])
    df1 = df1[['时间','CGI','eCGI','区县','区域维护部','厂家名称','所属E-NODEB','小区中文名',\
                                        '小区英文名','经度','纬度','设备维护状态','管理状态','覆盖类型','覆盖场景',\
                                        '上下行子帧配置','特殊子帧模式','小区带宽','工作频段','中心载频的信道号']]

    #进行dl/ulcomp、uplink_data_compression、dl 256qam、VMIMO统计
    print ('merging UlComp/DlComp/UDC/Dl 256qam...')
    comp = ul_dl_comp()
    df1 = pd.merge(df1, comp, left_on = 'CGI',right_on='CGI', how = 'left')
    print ('Finishing merging UlComp/DlComp/UDC/Dl 256qam...')

    #进行dlca统计
    print ('merging DlCa...')
    dlca = dl_ca()
    df1 = pd.merge(df1, dlca, left_on = 'CGI',right_on='CGI', how = 'left')
    print ('Finishing merging DlCa...')

    #进行ulca统计
    print ('merging UlCa...')
    ulca = ul_ca()
    df1 = pd.merge(df1, ulca, left_on = 'CGI',right_on='CGI', how = 'left')
    print ('Finishing merging UlCa...')
    
    #进行highspeedflag统计
    print ('merging HighSpeedFlag...')
    hs = high_speed()
    df1 = pd.merge(df1, hs, left_on = 'CGI',right_on='CGI', how = 'left')
    print ('Finishing merging HighSpeedFlag...')

    #进行ul 64qam统计
    print ('merging ul 64qam...')
    ulqam = ul_qam()

    df1 = pd.merge(df1, ulqam, left_on = 'CGI',right_on='CGI', how = 'left')
    print ('Finishing merging ul 64qam...')
    
    df1.to_csv('NewFunction0705.csv',index=True,header=True,encoding='gbk')

    #结束执行时间
    end = datetime.datetime.now()
    print (end-start)
