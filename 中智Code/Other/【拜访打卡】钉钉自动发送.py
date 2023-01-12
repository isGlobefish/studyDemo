# -*- coding: utf-8 -*-
'''
@createTool    : PyCharm-2020.3.2
@projectName   : pythonProjectPy3.9
@originalAuthor: Made in win10.Sys design by deHao.Zou
@createTime    : 2021/1/8 15:24
'''

import os
import re
import xlwt
import glob
import xlrd
import uuid
import time
import hmac
import json
import pyhdb
import base64
import shutil
import pymysql
import hashlib
import calendar
import requests
import datetime
import win32com
import pythoncom
import numpy as np
import urllib.parse
import pandas as pd
import urllib.request
from termcolor import cprint
from PIL import ImageGrab, Image
from time import strftime, gmtime
from qiniu import Auth, put_file, etag
from win32com.client import Dispatch, DispatchEx

# 定义时间
subTod7 = (datetime.datetime.now() + datetime.timedelta(days=-7)).strftime("%Y%m%d")  # 前七天
subTod1 = (datetime.datetime.now() + datetime.timedelta(days=-1)).strftime("%Y%m%d")  # 前一天
sub7Time = (datetime.datetime.now() + datetime.timedelta(days=-7)).strftime("%Y-%m-%d")  # 前七天
sub6Time = (datetime.datetime.now() + datetime.timedelta(days=-6)).strftime("%Y-%m-%d")  # 前六天
sub5Time = (datetime.datetime.now() + datetime.timedelta(days=-5)).strftime("%Y-%m-%d")  # 前无天
sub4Time = (datetime.datetime.now() + datetime.timedelta(days=-4)).strftime("%Y-%m-%d")  # 前四天
sub3Time = (datetime.datetime.now() + datetime.timedelta(days=-3)).strftime("%Y-%m-%d")  # 前三天
sub2Time = (datetime.datetime.now() + datetime.timedelta(days=-2)).strftime("%Y-%m-%d")  # 前两天
sub1Time = (datetime.datetime.now() + datetime.timedelta(days=-1)).strftime("%Y-%m-%d")  # 前一天

# 数据获取、可视化
try:
    print(">  获取HANA数据中,请稍等片刻")
    startGetHanaDataTime = datetime.datetime.now()


    # 获取 Connection 对象
    def get_HANA_Connection():
        connectionObj = pyhdb.connect(
            host="192.168.20.183",  # HANA地址
            port=30015,  # HANA端口号
            user="Hana106620",  # 用户名
            password="CHENjia90"  # 密码
        )
        return connectionObj


    # 获取拜访表指定时间段数据
    def get_matBF(connBF):
        cursorBF = connBF.cursor()
        cursorBF.execute(
            """SELECT * FROM "HD-HAND.SD.POWER_BI::CV_ZHRWQBF11" WHERE TO_DATE(CREATEDTIME)  >= '""" + subTod7 + """' AND TO_DATE(CREATEDTIME)  <= '""" + subTod1 + """'""")
        matBF = cursorBF.fetchall()
        return matBF


    # 获取签退表指定时间段数据
    def get_matQT(connQT):
        cursorQT = connQT.cursor()
        cursorQT.execute(
            """SELECT * FROM "HD-HAND.SD.POWER_BI::CV_ZHRWQBF12" WHERE TO_DATE(CREATEDTIME)  >= '""" + subTod7 + """' AND TO_DATE(CREATEDTIME)  <= '""" + subTod1 + """'""")
        matQT = cursorQT.fetchall()
        return matQT


    conn = get_HANA_Connection()
    dataBF = pd.DataFrame(get_matBF(conn))
    dataQT = pd.DataFrame(get_matQT(conn))
    endGetHanaDataTime = datetime.datetime.now()
    cprint(subTod7 + "-" + subTod1 + "数据成功获取, 拜访表数据有" + str(len(dataBF)) + "行; 签退表数据有" + str(len(dataQT)) + "行; 耗时：" + strftime(
        "%H:%M:%S", gmtime((endGetHanaDataTime - startGetHanaDataTime).seconds)), 'magenta', attrs=['bold', 'reverse', 'blink'])

except Exception as e:
    print("获取HANA数据出错！！！", e)

try:
    startMergeTableTime = datetime.datetime.now()

    dataBF.columns = ['MANDT_QD', 'OBJECTI_QD', 'NAME_QD', 'STAFF_DESC_QD', 'SY_DEPT_DESC_QD', 'SY_SFA_DESC_QD',
                      'SY_POINT_DESC_QD', 'F0000003_QD', 'SY_SFA_BINARYCODE_QD', 'SY_TYPE_QD', 'SY_SFA_ID_QD',
                      'SY_BFDATE_QD', 'key_QD', 'CREATEDBY_QD', 'SY_BFTIME_QD', 'MODIFIEDBY_QD', 'MODIFIEDTIME_QD',
                      'CREATEDBYOBJECT_QD', 'OWNERIDOBJECT_QD', 'OWNERDEPTIDOBJECT_QD', 'MODIFIEDBYOBJECT_QD',
                      'SY_LATITUDE_QD', 'SY_LONGITUDE_QD', 'STAFF_ID_QD', 'SY_DEPT_ID_QD']

    dataQT.columns = ['MANDT_QT', 'OBJECTID_QT', 'NAME_QT', 'STAFF_DESC_SECOND_QT', 'SY_DEPT_DESC_QT',
                      'SY_POINT_DESC_QT', 'SY_PHOTO_QT', 'SY_SFA_BINARYCODE_QT', 'SY_SFA_DESC_QT', 'SY_QTDATE_QT',
                      'XIAOJIE_QT', 'STAFF_DESC_FIRST_QT', 'SY_QTTIME_QT', 'MODIFIEDBY_QT', 'MODIFIEDTIME_QT',
                      'CREATEDBYOBJECT_QT', 'OWNERIDOBJECT_QT', 'OWNERDEPTIDOBJECT_QT', 'MODIFIEDBYOBJECT_QT',
                      'SY_LATITUDE_QT', 'SY_LONGITUDE_QT', 'STAFF_ID_QT', 'SY_DEPT_ID_QT', 'SY_SFA_ID_QT']

    dataBF['SY_BFDATE_QD'] = pd.to_datetime(dataBF['SY_BFDATE_QD'], format='%Y/%m/%d')
    dataBF['SY_BFDATE_QD'] = dataBF['SY_BFDATE_QD'].dt.date
    dataQT['SY_QTDATE_QT'] = pd.to_datetime(dataQT['SY_QTDATE_QT'], format='%Y/%m/%d')
    dataQT['SY_QTDATE_QT'] = dataQT['SY_QTDATE_QT'].dt.date

    dataBF['SY_BFTIME_QD'] = pd.to_datetime(dataBF['SY_BFTIME_QD'])
    dataQT['SY_QTTIME_QT'] = pd.to_datetime(dataQT['SY_QTTIME_QT'])

    '''
    数据预处理（清洗掉拜访时间、签退时间、人员编号和门店编号异常的数据）
    '''
    delVarListBF = ['SY_SFA_ID_QD', 'STAFF_ID_QD', 'SY_BFTIME_QD', 'SY_TYPE_QD']
    for delvarBF in delVarListBF:
        for i in range(len(dataBF[delvarBF])):
            if dataBF.loc[i, delvarBF] == 0 or dataBF.loc[i, delvarBF] == '':
                dataBF.loc[i, delvarBF] = np.nan
    dataBF = dataBF.dropna(axis=0, subset=delVarListBF)

    delVarListQT = ['SY_SFA_ID_QT', 'STAFF_ID_QT', 'SY_QTTIME_QT']
    for delvarQT in delVarListQT:
        for j in range(len(dataQT[delvarQT])):
            if dataQT.loc[j, delvarQT] == 0 or dataQT.loc[j, delvarQT] == '':
                dataQT.loc[j, delvarQT] = np.nan
    dataQT = dataQT.dropna(axis=0, subset=delVarListQT)

    df1 = dataBF.reset_index(drop=True)
    df2 = dataQT.reset_index(drop=True)

    df1.insert(0, 'indexQD', '')
    for i in range(len(df1['SY_BFTIME_QD'])):
        df1.loc[i, "indexQD"] = str(df1.loc[i, 'SY_BFTIME_QD'])[0:11] + str(df1.loc[i, 'STAFF_ID_QD']) + str(df1.loc[i, 'SY_SFA_ID_QD'])[0:9]

    df2.insert(0, 'indexQT', '')
    for j in range(len(df2['SY_QTTIME_QT'])):
        df2.loc[j, 'indexQT'] = str(df2.loc[j, 'SY_QTTIME_QT'])[0:11] + str(df2.loc[j, 'STAFF_ID_QT']) + str(df2.loc[j, 'SY_SFA_ID_QT'])[0:9]

    '''
    构造拜访表与签退表唯一的主键
    '''
    # 拜访表唯一主键构造
    startBFTime = datetime.datetime.now()
    print(">  拜访表主键唯一化")
    df1.insert(1, 'markQD', '')
    for i in range(len(df1['indexQD'])):
        list1 = np.where(df1.loc[:, 'indexQD'] == df1.loc[i, 'indexQD'])[0]
        list2 = np.argsort(list1)
        list_len = len(list1)
        arr_new = []
        for item in list1:
            arr_new.append(item)
        for item in list2:
            arr_new.append(item)
        for j in range(list_len):
            df1.loc[arr_new[j], 'markQD'] = "*" * arr_new[j + list_len]
    df1.insert(0, 'UNIQUE_KEYS', '')
    for i in range(len(df1['indexQD'])):
        df1.loc[i, 'UNIQUE_KEYS'] = str(df1.loc[i, 'indexQD']) + str(df1.loc[i, 'markQD'])
    endBFTime = datetime.datetime.now()
    print(">> 拜访表主键唯一化成功, 耗时：" + strftime("%H:%M:%S", gmtime((endBFTime - startBFTime).seconds)))

    # 签退表唯一主键
    startQTTime = datetime.datetime.now()
    print(">  签退表主键唯一化")
    df2.insert(0, 'UNIQUE_KEYS', '')
    for i in range(len(df2['indexQT'])):
        if len(np.where(df1.loc[:, 'indexQD'] == df2.loc[i, 'indexQT'])[0]) <= 1:
            if len(np.where(df1.loc[:, 'indexQD'] == df2.loc[i, 'indexQT'])[0]) == 0:
                df2.loc[i, 'UNIQUE_KEYS'] = df2.loc[i, 'indexQT']
            else:
                if len(np.where(df2.loc[:, 'indexQT'] == df2.loc[i, 'indexQT'])[0]) == 1:
                    df1_IndexOver = np.where(df1.loc[:, 'indexQD'] == df2.loc[i, 'indexQT'])[0]
                    end_TimeOver = time.mktime(
                        datetime.datetime.strptime(str(df2.loc[i, 'SY_QTTIME_QT']),
                                                   "%Y-%m-%d %H:%M:%S").timetuple())
                    start_TimeOver = time.mktime(
                        datetime.datetime.strptime(str(df1.loc[df1_IndexOver[0], 'SY_BFTIME_QD']),
                                                   "%Y-%m-%d %H:%M:%S").timetuple())
                    time_DiffOver = end_TimeOver - start_TimeOver
                    if time_DiffOver > 0:
                        df2.loc[i, 'UNIQUE_KEYS'] = df2.loc[i, 'indexQT']
                    else:
                        pass
                else:
                    if df2.loc[i, 'UNIQUE_KEYS'] != "":  # 之前填充过的唯一索引跳过
                        pass
                    else:
                        df1_IndexNew = np.where(df1.loc[:, 'indexQD'] == df2.loc[i, 'indexQT'])[0]
                        df2_IndexNew = np.where(df2.loc[:, 'indexQT'] == df2.loc[i, 'indexQT'])[0]
                        time_New1 = []
                        for ts1 in df1_IndexNew:
                            time_New1_Diff = df1.loc[ts1, 'SY_BFTIME_QD']
                            time_New1.append(time_New1_Diff)
                        time_New2 = []
                        for ts2 in df2_IndexNew:
                            time_New2_Diff = df2.loc[ts2, 'SY_QTTIME_QT']
                            time_New2.append(time_New2_Diff)
                        for tss in time_New1:
                            df1_NewList = [tss for i in time_New2]
                        new_Null = []
                        for s, ss in zip(time_New2, df1_NewList):
                            end_NewTime = time.mktime(
                                datetime.datetime.strptime(str(s), "%Y-%m-%d %H:%M:%S").timetuple())
                            start_NewTime = time.mktime(
                                datetime.datetime.strptime(str(ss), "%Y-%m-%d %H:%M:%S").timetuple())
                            time_NewDiff = end_NewTime - start_NewTime
                            new_Null.append(time_NewDiff)
                        # 全部小于0的情况跳过
                        if max([i for i in new_Null]) < 0:
                            pass
                        else:
                            # 定位出对应的大于等于0中最小的签到表索引
                            pos_NewNum_Min = min([i for i in new_Null if i >= 0])
                            for jj in range(len(new_Null)):
                                if new_Null[jj] == pos_NewNum_Min:
                                    loc_NewIndex = df2_IndexNew[jj]  # 正数且最小的索引
                                    df2.loc[loc_NewIndex, 'UNIQUE_KEYS'] = df1.loc[df1_IndexNew[0], 'UNIQUE_KEYS']
                                else:
                                    pass
        else:
            df1_Index = np.where(df1.loc[:, 'indexQD'] == df2.loc[i, 'indexQT'])[0]
            df2_Index = np.where(df2.loc[:, 'indexQT'] == df2.loc[i, 'indexQT'])[0]
            df1_Ser = np.argsort(df1_Index)
            df2_Ser = np.argsort(df2_Index)
            df1_len = len(df1_Ser)
            df2_len = len(df2_Ser)
            time_List1 = []
            for m in df1_Index:
                time_Diff1 = df1.loc[m, 'SY_BFTIME_QD']
                time_List1.append(time_Diff1)
            time_List2 = df2.loc[i, 'SY_QTTIME_QT']
            df2_Index_new = [time_List2 for i in time_List1]
            time_Null = []
            for k, kk in zip(df2_Index_new, time_List1):
                end_Time = time.mktime(datetime.datetime.strptime(str(k), "%Y-%m-%d %H:%M:%S").timetuple())
                start_Time = time.mktime(datetime.datetime.strptime(str(kk), "%Y-%m-%d %H:%M:%S").timetuple())
                time_Diff = end_Time - start_Time
                time_Null.append(time_Diff)
            # 全部小于0的情况跳过
            if max([i for i in time_Null]) < 0:
                pass
            else:
                # 定位出对应的大于等于0中最小的签到表索引
                pos_Num_Min = min([i for i in time_Null if i >= 0])
                for jj in range(len(time_Null)):
                    if time_Null[jj] == pos_Num_Min:
                        loc_Index = df1_Index[jj]  # 正数且最小的索引
                        df2.loc[i, 'UNIQUE_KEYS'] = df1.loc[loc_Index, 'UNIQUE_KEYS']
    endQTTime = datetime.datetime.now()
    print(">> 签退表主键唯一化成功, 耗时：" + strftime("%H:%M:%S", gmtime((endQTTime - startQTTime).seconds)))

    result = pd.merge(df1.drop_duplicates(), df2.drop_duplicates(), how='left', left_on='UNIQUE_KEYS', right_on='UNIQUE_KEYS')
    # 多级排序，ascending=False代表按降序排序，na_position='last'代表空值放在最后一位
    result.sort_values(by=['UNIQUE_KEYS', 'SY_QTTIME_QT'], ascending=False, na_position='last')
    result.drop_duplicates(subset='UNIQUE_KEYS', keep='last', inplace=True)
    # res = result.drop(columns=['indexQD', 'indexQT', 'markQD'])
    res = result.drop(columns=['markQD', 'indexQT'])
    # 特殊处理个别
    # res = res[res['OBJECTI_QD'] != 'd50d3279-34d3-4cd8-8a57-c0b7bba143d6']
    # res.to_excel('D:/DataCenter/VisImage/dropDate.xlsx')
    origData = res.reset_index(drop=True)
    # origData.to_excel('C:/Users/Zeus/Desktop/近七天拜访签退.xlsx', index=False)
    endMergeTableTime = datetime.datetime.now()
    cprint(subTod7 + "-" + subTod1 + "拜访签退表合并成功, 匹配成功" + str(len(res)) + "行; 耗时：" + strftime("%H:%M:%S", gmtime(
        (endMergeTableTime - startMergeTableTime).seconds)), 'magenta', attrs=['bold', 'reverse', 'blink'])
except Exception as e:
    print("数据处理or表格合并出错！！！", e)

try:
    # 打卡记录
    # origData = pd.read_excel('D:/DataCenter/VisImage/近七天拜访签退.xlsx')
    origTable = pd.pivot_table(origData, index=['CREATEDBY_QD', 'SY_TYPE_QD'], columns=['SY_BFDATE_QD'], values=['UNIQUE_KEYS'],
                               aggfunc={'UNIQUE_KEYS': np.count_nonzero}, fill_value=0, margins=True, margins_name='总计')
    origTable.to_excel('D:/DataCenter/VisImage/orig123.xlsx')
    origDF = pd.read_excel('D:/DataCenter/VisImage/orig123.xlsx', header=0, skiprows=[1, 2])
    origDF.columns = ['Name', 'Type', 'daySub7', 'daySub6', 'daySub5', 'daySub4', 'daySub3', 'daySub2', 'daySub1', 'sumRecord']

    # 拜访门店数
    dealData = origData.drop_duplicates(['indexQD']).reset_index(drop=True)
    dealTable = pd.pivot_table(dealData, index=['CREATEDBY_QD', 'SY_TYPE_QD'], columns=['SY_BFDATE_QD'], values=['indexQD'],
                               aggfunc={'indexQD': np.count_nonzero}, fill_value=0, margins=True, margins_name='总计')
    dealTable.to_excel('D:/DataCenter/VisImage/deal123.xlsx')
    dealDF = pd.read_excel('D:/DataCenter/VisImage/deal123.xlsx', header=0, skiprows=[1, 2])
    dealDF.columns = ['Name', 'Type', 'daySub7', 'daySub6', 'daySub5', 'daySub4', 'daySub3', 'daySub2', 'daySub1', 'sumSFA']


    def newDF():
        dfNew = pd.DataFrame(
            columns=['Name', 'Type', 'sumRecord', 'sumSFA', 'recordDaySub7', 'recordDaySub6', 'recordDaySub5', 'recordDaySub4', 'recordDaySub3',
                     'recordDaySub2', 'recordDaySub1', 'SFADaySub7', 'SFADaySub6', 'SFADaySub5', 'SFADaySub4', 'SFADaySub3', 'SFADaySub2',
                     'SFADaySub1', 'partner'])
        return dfNew


    # 字典匹配
    sumSFADict = {}
    SFADaySub7Dict = {}
    SFADaySub6Dict = {}
    SFADaySub5Dict = {}
    SFADaySub4Dict = {}
    SFADaySub3Dict = {}
    SFADaySub2Dict = {}
    SFADaySub1Dict = {}
    partnerDict = {}  # 伙伴
    sfaRow = len(dealDF)
    for i in range(0, sfaRow):
        sumSFADict[dealDF.loc[i, 'Name']] = dealDF.loc[i, 'sumSFA']
        SFADaySub7Dict[dealDF.loc[i, 'Name']] = dealDF.loc[i, 'daySub7']
        SFADaySub6Dict[dealDF.loc[i, 'Name']] = dealDF.loc[i, 'daySub6']
        SFADaySub5Dict[dealDF.loc[i, 'Name']] = dealDF.loc[i, 'daySub5']
        SFADaySub4Dict[dealDF.loc[i, 'Name']] = dealDF.loc[i, 'daySub4']
        SFADaySub3Dict[dealDF.loc[i, 'Name']] = dealDF.loc[i, 'daySub3']
        SFADaySub2Dict[dealDF.loc[i, 'Name']] = dealDF.loc[i, 'daySub2']
        SFADaySub1Dict[dealDF.loc[i, 'Name']] = dealDF.loc[i, 'daySub1']
    dictionary = xlrd.open_workbook('D:/DataCenter/VisImage/伙伴对照表.xlsx')
    sheet = dictionary.sheet_by_name('Sheet1')
    partnerRow = sheet.nrows
    for i in range(1, partnerRow):
        values = sheet.row_values(i)
        partnerDict[values[1]] = values[2]

    readyTable = newDF()
    print(">  数据表归一化")
    readyTable['Name'] = origDF['Name']
    readyTable['Type'] = origDF['Type']
    readyTable['sumRecord'] = origDF['sumRecord']
    readyTable['sumSFA'] = readyTable.apply(lambda x: sumSFADict.setdefault(x['Name'], 0), axis=1)
    readyTable['recordDaySub7'] = origDF['daySub7']
    readyTable['recordDaySub6'] = origDF['daySub6']
    readyTable['recordDaySub5'] = origDF['daySub5']
    readyTable['recordDaySub4'] = origDF['daySub4']
    readyTable['recordDaySub3'] = origDF['daySub3']
    readyTable['recordDaySub2'] = origDF['daySub2']
    readyTable['recordDaySub1'] = origDF['daySub1']
    readyTable['SFADaySub7'] = readyTable.apply(lambda x: SFADaySub7Dict.setdefault(x['Name'], 0), axis=1)
    readyTable['SFADaySub6'] = readyTable.apply(lambda x: SFADaySub6Dict.setdefault(x['Name'], 0), axis=1)
    readyTable['SFADaySub5'] = readyTable.apply(lambda x: SFADaySub5Dict.setdefault(x['Name'], 0), axis=1)
    readyTable['SFADaySub4'] = readyTable.apply(lambda x: SFADaySub4Dict.setdefault(x['Name'], 0), axis=1)
    readyTable['SFADaySub3'] = readyTable.apply(lambda x: SFADaySub3Dict.setdefault(x['Name'], 0), axis=1)
    readyTable['SFADaySub2'] = readyTable.apply(lambda x: SFADaySub2Dict.setdefault(x['Name'], 0), axis=1)
    readyTable['SFADaySub1'] = readyTable.apply(lambda x: SFADaySub1Dict.setdefault(x['Name'], 0), axis=1)
    readyTable['partner'] = readyTable.apply(lambda x: partnerDict.setdefault(x['Name'], 0), axis=1)
    # 多级排序，ascending=False代表按降序排序，na_position='last'代表空值放在最后一位
    readyExport = readyTable.sort_values(by=['sumRecord', 'sumSFA'], ascending=False, na_position='last')
    # readyExport.to_excel('C:/Users/Zeus/Desktop/5555.xlsx')
except Exception as e:
    print("数据处理or表格合并出错！！！", e)


# 制作表格
def Create_TableImage(Ptable, FilesName):
    # 设置单元格样式
    def set_style(fontName, height, bold=False, Halign=False, Valign=False, setBorder=False, setbgcolor=False):
        style = xlwt.XFStyle()  # 设置类型

        font = xlwt.Font()  # 为样式创建字体
        font.name = fontName
        font.height = height  # 字体大小，220就是11号字体，大概就是11*20得来
        font.bold = bold
        font.color = 'black'
        font.color_index = 4
        style.font = font

        alignment = xlwt.Alignment()  # 设置字体在单元格的位置
        if Halign == 0:
            alignment.horz = xlwt.Alignment.HORZ_CENTER  # 水平居中
        elif Halign == 1:
            alignment.horz = xlwt.Alignment.HORZ_LEFT  # 水平偏左
        else:
            alignment.horz = xlwt.Alignment.HORZ_RIGHT  # 水平偏右
        if Valign == 0:
            alignment.vert = xlwt.Alignment.VERT_CENTER  # 竖直居中
        elif Valign == 1:
            alignment.vert = xlwt.Alignment.VERT_TOP  # 竖直置顶
        else:
            alignment.vert = xlwt.Alignment.VERT_BOTTOM  # 竖直底部
        # alignment.horz = xlwt.Alignment.HORZ_CENTER  # 水平居中
        # alignment.horz = xlwt.Alignment.HORZ_LEFT  # 水平偏左
        # alignment.horz = xlwt.Alignment.HORZ_RIGHT  # 水平偏右
        # alignment.vert = xlwt.Alignment.VERT_CENTER  # 竖直居中
        # alignment.vert = xlwt.Alignment.VERT_TOP  # 竖直置顶
        # alignment.vert = xlwt.Alignment.VERT_BOTTOM  # 竖直底部
        style.alignment = alignment

        border = xlwt.Borders()  # 给单元格加框线
        if setBorder == 0:
            border.left = xlwt.Borders.THIN  # 左
            border.top = xlwt.Borders.THIN  # 上
            border.right = xlwt.Borders.THIN  # 右
            border.bottom = xlwt.Borders.THIN  # 下
            border.left_colour = 0x40  # 设置框线颜色，0x40是黑色，颜色真的巨多
            border.right_colour = 0x40
            border.top_colour = 0x40
            border.bottom_colour = 0x40
        else:
            pass
        style.borders = border

        pattern = xlwt.Pattern()  # 设置背景颜色
        if setbgcolor == 0:
            pattern.pattern = xlwt.Pattern.SOLID_PATTERN
            pattern.pattern_fore_colour = xlwt.Style.colour_map['turquoise']
        else:
            pass
        style.pattern = pattern
        return style

    try:
        # 创建表格模板并写入数据
        f = xlwt.Workbook()  # 创建工作簿
        sheet1 = f.add_sheet('Sheet1', cell_overwrite_ok=True)

        # 设置列宽
        for i in range(18):
            if i == 0:
                sheet1.col(i).width = 256 * 15
            elif i == 1:
                sheet1.col(i).width = 256 * 12
            elif i == 2 or i == 3:
                sheet1.col(i).width = 256 * 14
            else:
                sheet1.col(i).width = 256 * 8

        # 设置行高
        for rowi in range(0, 3):
            rowHeight = xlwt.easyxf('font:height 220;')  # 36pt,类型小初的字号
            rowNum = sheet1.row(rowi)
            rowNum.set_style(rowHeight)

        """ 设计表头 """
        # 第A列
        sheet1.write_merge(0, 2, 0, 0, "姓名",
                           set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
        # 第B列
        sheet1.write_merge(0, 2, 1, 1, "打卡类型",
                           set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
        # 第C-D列
        sheet1.write_merge(0, 0, 2, 3, sub7Time + " 至 " + sub1Time,
                           set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
        # 第C列
        sheet1.write_merge(1, 2, 2, 2, "总打卡次数/条",
                           set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
        # 第D列
        sheet1.write_merge(1, 2, 3, 3, "总拜访门店/家",
                           set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
        # 第E-R列
        sheet1.write_merge(0, 0, 4, 17, "近七天拜访打卡记录",
                           set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
        # 第E-K列
        sheet1.write_merge(1, 1, 4, 10, "每日打卡记录次数/条",
                           set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
        # 第L-R列
        sheet1.write_merge(1, 1, 11, 17, "每日拜访门店数/家",
                           set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
        # 第E列
        sheet1.write_merge(2, 2, 4, 4, str(int(sub7Time.split('-')[1])) + "月" + str(int(sub7Time.split('-')[2])) + "日",
                           set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
        # 第F列
        sheet1.write_merge(2, 2, 5, 5, str(int(sub6Time.split('-')[1])) + "月" + str(int(sub6Time.split('-')[2])) + "日",
                           set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
        # 第G列
        sheet1.write_merge(2, 2, 6, 6, str(int(sub5Time.split('-')[1])) + "月" + str(int(sub5Time.split('-')[2])) + "日",
                           set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
        # 第H列
        sheet1.write_merge(2, 2, 7, 7, str(int(sub4Time.split('-')[1])) + "月" + str(int(sub4Time.split('-')[2])) + "日",
                           set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
        # 第I列
        sheet1.write_merge(2, 2, 8, 8, str(int(sub3Time.split('-')[1])) + "月" + str(int(sub3Time.split('-')[2])) + "日",
                           set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
        # 第J列
        sheet1.write_merge(2, 2, 9, 9, str(int(sub2Time.split('-')[1])) + "月" + str(int(sub2Time.split('-')[2])) + "日",
                           set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
        # 第K列
        sheet1.write_merge(2, 2, 10, 10, str(int(sub1Time.split('-')[1])) + "月" + str(int(sub1Time.split('-')[2])) + "日",
                           set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
        # 第L列
        sheet1.write_merge(2, 2, 11, 11, str(int(sub7Time.split('-')[1])) + "月" + str(int(sub7Time.split('-')[2])) + "日",
                           set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
        # 第M列
        sheet1.write_merge(2, 2, 12, 12, str(int(sub6Time.split('-')[1])) + "月" + str(int(sub6Time.split('-')[2])) + "日",
                           set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
        # 第N列
        sheet1.write_merge(2, 2, 13, 13, str(int(sub5Time.split('-')[1])) + "月" + str(int(sub5Time.split('-')[2])) + "日",
                           set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
        # 第O列
        sheet1.write_merge(2, 2, 14, 14, str(int(sub4Time.split('-')[1])) + "月" + str(int(sub4Time.split('-')[2])) + "日",
                           set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
        # 第P列
        sheet1.write_merge(2, 2, 15, 15, str(int(sub3Time.split('-')[1])) + "月" + str(int(sub3Time.split('-')[2])) + "日",
                           set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
        # 第Q列
        sheet1.write_merge(2, 2, 16, 16, str(int(sub2Time.split('-')[1])) + "月" + str(int(sub2Time.split('-')[2])) + "日",
                           set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
        # 第R列
        sheet1.write_merge(2, 2, 17, 17, str(int(sub1Time.split('-')[1])) + "月" + str(int(sub1Time.split('-')[2])) + "日",
                           set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))

        rowLen = len(Ptable)  # 行数
        sendColumns = ['Name', 'Type', 'sumRecord', 'sumSFA', 'recordDaySub7', 'recordDaySub6', 'recordDaySub5', 'recordDaySub4', 'recordDaySub3',
                       'recordDaySub2', 'recordDaySub1', 'SFADaySub7', 'SFADaySub6', 'SFADaySub5', 'SFADaySub4', 'SFADaySub3', 'SFADaySub2',
                       'SFADaySub1']
        # 填充表格数据
        for rowi in range(rowLen):
            for colj, colName in enumerate(sendColumns):
                sheet1.write_merge(rowi + 3, rowi + 3, colj, colj, str(Ptable.loc[rowi, colName]),
                                   set_style('等线', 210, False, Halign=0, Valign=0, setBorder=0, setbgcolor=1))

        # for rowi in range(rowLen):
        #     sheet1.write_merge(rowi + 3, rowi + 3, 0, 0, str(Ptable.loc[rowi, 'Name']),
        #                        set_style('等线', 210, False, Halign=0, Valign=0, setBorder=0, setbgcolor=1))
        #     sheet1.write_merge(rowi + 3, rowi + 3, 1, 1, str(Ptable.loc[rowi, 'Type']),
        #                        set_style('等线', 210, False, Halign=0, Valign=0, setBorder=0, setbgcolor=1))
        #     sheet1.write_merge(rowi + 3, rowi + 3, 2, 2, str(Ptable.loc[rowi, 'sumRecord']),
        #                        set_style('等线', 210, False, Halign=0, Valign=0, setBorder=0, setbgcolor=1))
        #     sheet1.write_merge(rowi + 3, rowi + 3, 3, 3, str(Ptable.loc[rowi, 'sumSFA']),
        #                        set_style('等线', 210, False, Halign=0, Valign=0, setBorder=0, setbgcolor=1))
        #     sheet1.write_merge(rowi + 3, rowi + 3, 4, 4, str(Ptable.loc[rowi, 'recordDaySub7']),
        #                        set_style('等线', 210, False, Halign=0, Valign=0, setBorder=0, setbgcolor=1))
        #     sheet1.write_merge(rowi + 3, rowi + 3, 5, 5, str(Ptable.loc[rowi, 'recordDaySub6']),
        #                        set_style('等线', 210, False, Halign=0, Valign=0, setBorder=0, setbgcolor=1))
        #     sheet1.write_merge(rowi + 3, rowi + 3, 6, 6, str(Ptable.loc[rowi, 'recordDaySub5']),
        #                        set_style('等线', 210, False, Halign=0, Valign=0, setBorder=0, setbgcolor=1))
        #     sheet1.write_merge(rowi + 3, rowi + 3, 7, 7, str(Ptable.loc[rowi, 'recordDaySub4']),
        #                        set_style('等线', 210, False, Halign=0, Valign=0, setBorder=0, setbgcolor=1))
        #     sheet1.write_merge(rowi + 3, rowi + 3, 8, 8, str(Ptable.loc[rowi, 'recordDaySub3']),
        #                        set_style('等线', 210, False, Halign=0, Valign=0, setBorder=0, setbgcolor=1))
        #     sheet1.write_merge(rowi + 3, rowi + 3, 9, 9, str(Ptable.loc[rowi, 'recordDaySub2']),
        #                        set_style('等线', 210, False, Halign=0, Valign=0, setBorder=0, setbgcolor=1))
        #     sheet1.write_merge(rowi + 3, rowi + 3, 10, 10, str(Ptable.loc[rowi, 'recordDaySub1']),
        #                        set_style('等线', 210, False, Halign=0, Valign=0, setBorder=0, setbgcolor=1))
        #     sheet1.write_merge(rowi + 3, rowi + 3, 11, 11, str(Ptable.loc[rowi, 'SFADaySub7']),
        #                        set_style('等线', 210, False, Halign=0, Valign=0, setBorder=0, setbgcolor=1))
        #     sheet1.write_merge(rowi + 3, rowi + 3, 12, 12, str(Ptable.loc[rowi, 'SFADaySub6']),
        #                        set_style('等线', 210, False, Halign=0, Valign=0, setBorder=0, setbgcolor=1))
        #     sheet1.write_merge(rowi + 3, rowi + 3, 13, 13, str(Ptable.loc[rowi, 'SFADaySub5']),
        #                        set_style('等线', 210, False, Halign=0, Valign=0, setBorder=0, setbgcolor=1))
        #     sheet1.write_merge(rowi + 3, rowi + 3, 14, 14, str(Ptable.loc[rowi, 'SFADaySub4']),
        #                        set_style('等线', 210, False, Halign=0, Valign=0, setBorder=0, setbgcolor=1))
        #     sheet1.write_merge(rowi + 3, rowi + 3, 15, 15, str(Ptable.loc[rowi, 'SFADaySub3']),
        #                        set_style('等线', 210, False, Halign=0, Valign=0, setBorder=0, setbgcolor=1))
        #     sheet1.write_merge(rowi + 3, rowi + 3, 16, 16, str(Ptable.loc[rowi, 'SFADaySub2']),
        #                        set_style('等线', 210, False, Halign=0, Valign=0, setBorder=0, setbgcolor=1))
        #     sheet1.write_merge(rowi + 3, rowi + 3, 17, 17, str(Ptable.loc[rowi, 'SFADaySub1']),
        #                        set_style('等线', 210, False, Halign=0, Valign=0, setBorder=0, setbgcolor=1))

        # 添加求和行
        sheet1.write_merge(rowLen + 3, rowLen + 3, 0, 1, "求和",
                           set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
        sheet1.write_merge(rowLen + 3, rowLen + 3, 2, 2, str(sum(Ptable.loc[:, 'sumRecord'])),
                           set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
        sheet1.write_merge(rowLen + 3, rowLen + 3, 3, 3, str(sum(Ptable.loc[:, 'sumSFA'])),
                           set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
        sheet1.write_merge(rowLen + 3, rowLen + 3, 4, 4, str(sum(Ptable.loc[:, 'recordDaySub7'])),
                           set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
        sheet1.write_merge(rowLen + 3, rowLen + 3, 5, 5, str(sum(Ptable.loc[:, 'recordDaySub6'])),
                           set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
        sheet1.write_merge(rowLen + 3, rowLen + 3, 6, 6, str(sum(Ptable.loc[:, 'recordDaySub5'])),
                           set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
        sheet1.write_merge(rowLen + 3, rowLen + 3, 7, 7, str(sum(Ptable.loc[:, 'recordDaySub4'])),
                           set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
        sheet1.write_merge(rowLen + 3, rowLen + 3, 8, 8, str(sum(Ptable.loc[:, 'recordDaySub3'])),
                           set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
        sheet1.write_merge(rowLen + 3, rowLen + 3, 9, 9, str(sum(Ptable.loc[:, 'recordDaySub2'])),
                           set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
        sheet1.write_merge(rowLen + 3, rowLen + 3, 10, 10, str(sum(Ptable.loc[:, 'recordDaySub1'])),
                           set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
        sheet1.write_merge(rowLen + 3, rowLen + 3, 11, 11, str(sum(Ptable.loc[:, 'SFADaySub7'])),
                           set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
        sheet1.write_merge(rowLen + 3, rowLen + 3, 12, 12, str(sum(Ptable.loc[:, 'SFADaySub6'])),
                           set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
        sheet1.write_merge(rowLen + 3, rowLen + 3, 13, 13, str(sum(Ptable.loc[:, 'SFADaySub5'])),
                           set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
        sheet1.write_merge(rowLen + 3, rowLen + 3, 14, 14, str(sum(Ptable.loc[:, 'SFADaySub4'])),
                           set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
        sheet1.write_merge(rowLen + 3, rowLen + 3, 15, 15, str(sum(Ptable.loc[:, 'SFADaySub3'])),
                           set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
        sheet1.write_merge(rowLen + 3, rowLen + 3, 16, 16, str(sum(Ptable.loc[:, 'SFADaySub2'])),
                           set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
        sheet1.write_merge(rowLen + 3, rowLen + 3, 17, 17, str(sum(Ptable.loc[:, 'SFADaySub1'])),
                           set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))

    except Exception as e:
        print("制作发送表格出错！！！")

    f.save('D:/DataCenter/VisImage/Files/' + str(FilesName) + '.xls')  # 保存文件

    try:
        # screenArea——格式类似"A1:J10"
        def excelCatchScreen(file_name, sheet_name, screen_area, save_path, img_name=False):
            pythoncom.CoInitialize()  # excel多线程相关
            Application = win32com.client.gencache.EnsureDispatch("Excel.Application")  # 启动excel
            Application.Visible = False  # 可视化
            Application.DisplayAlerts = False  # 是否显示警告
            wb = Application.Workbooks.Open(file_name)  # 打开excel
            ws = wb.Sheets(sheet_name)  # 选择Sheet
            ws.Range(screen_area).CopyPicture()  # 复制图片区域
            time.sleep(1)
            ws.Paste()  # 粘贴 ws.Paste(ws.Range('B1'))  # 将图片移动到具体位置

            # name = str(uuid.uuid4())  # 重命名唯一值
            name = str(FilesName)
            new_shape_name = name[:6]
            Application.Selection.ShapeRange.Name = new_shape_name  # 将刚刚选择的Shape重命名, 避免与已有图片混淆
            ws.Shapes(new_shape_name).Copy()  # 选择图片
            time.sleep(1)
            img = ImageGrab.grabclipboard()  # 获取剪贴板的图片数据
            if not img_name:
                img_name = name + ".PNG"
            img.save(save_path + img_name)  # 保存图片
            time.sleep(1)
            # wb.Save()
            time.sleep(1)
            wb.Close(SaveChanges=0)  # 关闭工作薄，不保存
            time.sleep(1)
            Application.Quit()  # 退出excel
            pythoncom.CoUninitialize()

        excelCatchScreen('D:/DataCenter/VisImage/Files/' + str(FilesName) + '.xls', "Sheet1", "A1:R" + str(rowLen + 4),
                         'D:/DataCenter/VisImage/Image/')

    except Exception as e:
        print('截取图片出错！！！')


# 上传本地图片获取网上图片URL
def get_image_url(imagePath):
    if str(imagePath).split('.')[-1] == 'jpg' or str(imagePath).split('.')[-1] == 'JPG':
        key = 'transitPic_abc' + '.' + str(imagePath).split('.')[-1]  # 七牛云网盘文件名
    elif str(imagePath).split('.')[-1] == 'png' or str(imagePath).split('.')[-1] == 'PNG':
        key = 'transitPic_123' + '.' + str(imagePath).split('.')[-1]  # 七牛云网盘文件名
    else:
        print("请检查图片格式！！！")
    # 七牛云密钥管理：https://portal.qiniu.com/user/key
    # 【账号：144714959@qq.com  密码：thebtx1997】
    access_key = "DZnCErimkn2yQrn4aYel3JX7vPXKRonlvDFoVh1e"
    secret_key = "FBEHIFyMG28nWZrn316df-ny5bmIz_LanRWtabCi"
    q = Auth(access_key, secret_key)
    bucket_name = "qiniu730173201"  # 七牛云盘名
    token = q.upload_token(bucket_name, key)  # 删掉旧图片
    time.sleep(3)
    ret, info = put_file(token, key, imagePath)  # 上传新图片
    time.sleep(3)
    baseURL = "http://zzsy.zeus.cn/"  # 中智二级域名
    subURL = baseURL + '/' + key
    time.sleep(4)  # 等待4秒(服务器在上海,存在延迟)
    pictureURL = q.private_download_url(subURL)  # 链接图片URL
    return pictureURL


# 定义钉钉功能
class dingdingFunction(object):
    def __init__(self, roboturl, robotsecret, appkey, appsecret):
        """
        :param roboturl: 群机器人WebHook_url
        :param robotsecret: 安全设置的加签秘钥
        :param appkey: 企业开发平台小程序AppKey
        :param appsecret: 企业开发平台小程序AppSecret
        """
        self.roboturl = roboturl
        self.robotsecret = robotsecret
        self.appkey = appkey
        self.appsecret = appsecret
        timestamp = round(time.time() * 1000)  # 时间戳
        secret_enc = robotsecret.encode('utf-8')
        string_to_sign = '{}\n{}'.format(timestamp, robotsecret)
        string_to_sign_enc = string_to_sign.encode('utf-8')
        hmac_code = hmac.new(secret_enc, string_to_sign_enc, digestmod=hashlib.sha256).digest()
        sign = urllib.parse.quote_plus(base64.b64encode(hmac_code))  # 最终签名
        self.webhook_url = self.roboturl + '&timestamp={}&sign={}'.format(timestamp, sign)  # 最终url,url+时间戳+签名

    # 发送文件
    def getAccess_token(self):
        url = 'https://oapi.dingtalk.com/gettoken?appkey=%s&appsecret=%s' % (AppKey, AppSecret)
        headers = {
            'Content-Type': "application/x-www-form-urlencoded"
        }
        data = {'appkey': self.appkey,
                'appsecret': self.appsecret}
        r = requests.request('GET', url, data=data, headers=headers)
        access_token = r.json()["access_token"]
        return access_token

    def getMedia_id(self, filespath):
        access_token = self.getAccess_token()  # 拿到接口凭证
        url = 'https://oapi.dingtalk.com/media/upload?access_token=' + access_token + '&type=file'
        files = {'media': open(filespath, 'rb')}
        data = {'access_token': access_token,
                'type': 'file'}
        response = requests.post(url, files=files, data=data)
        json = response.json()
        return json["media_id"]

    def sendFile(self, chatid, filespath):
        access_token = self.getAccess_token()
        media_id = self.getMedia_id(filespath)
        url = 'https://oapi.dingtalk.com/chat/send?access_token=' + access_token
        header = {
            'Content-Type': 'application/json'
        }
        data = {'access_token': access_token,
                'chatid': chatid,
                'msg': {
                    'msgtype': 'file',
                    'file': {'media_id': media_id}
                }}
        r = requests.request('POST', url, data=json.dumps(data), headers=header)
        print(r.json()["errmsg"])

    # 发送消息
    def sendMessage(self, content, chatName, num, sum):
        """
        :param content: 发送内容
        """
        header = {
            "Content-Type": "application/json",
            "Charset": "UTF-8"
        }
        sendContent = json.dumps(content)  # 将字典类型数据转化为json格式
        sendContent = sendContent.encode("utf-8")  # 编码为UTF-8格式
        request = urllib.request.Request(url=self.webhook_url, data=sendContent, headers=header)  # 发送请求
        opener = urllib.request.urlopen(request)  # 将请求发回的数据构建成为文件格式
        print('>>> ' + str(num).zfill(3) + '/' + str(sum).zfill(3) + ' ' + chatName)  # 返回发送结果


#     加一个发送文件（包括图片、文本、表格、压缩文件等等）
#     获取全部手机号（匹配人名与手机号）


if __name__ == '__main__':
    orgChatData = pd.read_excel('D:/DataCenter/VisImage/伙伴对照表.xlsx', header=0)
    # orgChatData = pd.read_excel('C:/Users/Zeus/Desktop/伙伴对照表.xlsx', header=0)
    ChatData = orgChatData.drop_duplicates(['Partner']).reset_index(drop=True)

    AppKey = 'dingjpjkc2vaqjoqgmhz'  # 企业开发平台小程序AppKey
    AppSecret = 'oKNcuSF12oW0j9eBeO53wA6qwmKCVz34NVy1NvtvnjsvKPOdKiozsSZzUypNSWDc'  # 企业开发平台小程序AppSecret

    for ichat in range(len(ChatData)):
        partnerTable = readyExport[readyExport['partner'] == str(ChatData.loc[ichat, 'Partner'])].reset_index(drop=True)
        # partnerTable = readyExport[readyExport['partner'] == "广韶清中珠_郑路线"].reset_index(drop=True)
        try:
            Create_TableImage(partnerTable, str(ChatData.loc[ichat, 'Partner']))  # 创建表格、制作发送图片
            # Create_TableImage(partnerTable, "广韶清中珠_郑路线")  # 创建表格、制作发送图片

            ddMessage = {  # 发布消息内容
                "msgtype": "markdown",
                "markdown": {"title": "草晶华每日门店打卡数据",  # @某人 才会显示标题
                             "text": "\n> 大家好！我是 数运小助手 机器人。 **" + sub7Time + "** 至 **" + sub1Time +
                                     "** 近七天的门店打卡数据已经送达, 请大家注意查看哦(⊙o⊙), 如有疑问请及时回馈, 感谢大家的理解与支持。"
                                     "\n![Image](" + get_image_url(
                                 'D:/DataCenter/VisImage/Image/' + str(ChatData.loc[ichat, 'Partner']) + '.PNG') + ")"
                                                                                                                   "\n###### ----------------------------------------------"
                                                                                                                   "\n###### 发布时间：" +
                                     str(datetime.datetime.now()).split('.')[0]},  # 发布时间
                "at": {
                    # "atMobiles": [15817552982],  # 指定@某人
                    "isAtAll": False  # 是否@所有人[False:否, True:是]
                }
            }

            RobotWebHookURL = ChatData.loc[ichat, 'RobotURL']  # 群机器人url
            RobotSecret = ChatData.loc[ichat, 'RobotSecret']  # 群机器人加签秘钥secret(默认数运小跑腿)

            """
            特别说明：
                    发送消息：目前支持text、link、markdown等形式文字及图片，新增支持本地文件和图片类媒体文件的发送.
                    发送文件：目前支持简单excel表(csv、xlsx、xls等)、word、压缩文件,不支持ppt等文件的发送.
            """

            # 发送消息
            dingdingFunction(RobotWebHookURL, RobotSecret, AppKey, AppSecret).sendMessage(ddMessage, ChatData.loc[ichat, 'ChatName'], ichat + 1,
                                                                                          len(ChatData))
        except:
            pass

        # partnerTable.to_excel('C:/Users/Zeus/Desktop/123.xlsx',index=False)
