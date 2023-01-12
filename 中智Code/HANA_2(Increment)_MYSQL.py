# -*- coding: utf-8 -*-
'''
@createTool    : PyCharm-2020.2.2
@projectName   : pythonProject
@originalAuthor: Made in win10.Sys design by deHao.Zou
@createTime    : 2020/10/28 10:15
'''
import glob
import time
import pymysql
import datetime
import numpy as np
import pandas as pd
from termcolor import cprint
from time import strftime, gmtime
import pyhdb  # 加载连接HANA的所需模块
from dateutil.relativedelta import relativedelta
from sqlalchemy import create_engine  # 连接mysql模块
from sqlalchemy.types import NVARCHAR

"""
=======================================================================
HANA信息详情                                                           #
=======================================================================
HANA地址：192.168.20.183                                               #
HANA端口：30015                                                        #
HANA名：Hana106620                                                     #
HANA密码：CHENjia90                                                    #
拜访表名：HBP(HANA106620).Content.HD-HAND.SD.POWER_BI.CV_ZHRWQBF11      #
离店表名：HBP(HANA106620).Content.HD-HAND.SD.POWER_BI.CV_ZHRWQBF12      #
------------------------------------------------------------------------
------------------------------------------------------------------------
------------------------------------------------------------------------
========================================================================
MySQL信息详情                                                           #
========================================================================
MySQL地址：192.168.20228                                                #
MySQL端口：3306                                                         #
MySQL名(一般默认): root                                                 #
MySQL密码：123456                                                       #
数据上传位置：alex -> sfa -> sfa_bf_fact                                 #
------------------------------------------------------------------------
需求：
从HANA里面提取拜访表与签退表两张表数据, 合并两张表（匹配拜访时间与签退时间）,把处理好的表上传数据库保存。

"""

startTime = datetime.datetime.now()


# 定义上传数据库对象
def uploadSQL(uploadTable, targetDB, updateDBTable):
    '''
    uploadTable   :待上传表格
    TargetDB      ：目标数据库
    updateDBTable ：目标数据库下的表格
    '''
    try:
        # 转换数据表格式
        def mapping_df_types(convTabFormat):
            dtypedict = {}
            for i, j in zip(convTabFormat.columns, convTabFormat.dtypes):
                if "object" in str(j):
                    dtypedict.update({i: NVARCHAR(length=255)})
                if "float" in str(j):
                    dtypedict.update({i: NVARCHAR(length=255)})
                if "int" in str(j):
                    dtypedict.update({i: NVARCHAR(length=255)})
    except Exception as e:
        print("数据格式约束出错！！！", e)
    try:
        dtypedict = mapping_df_types(uploadTable)  # 转换数据格式
        engine = create_engine("mysql+pymysql://root:123456@localhost:3306/" + str(targetDB), encoding='utf-8',
                               echo=False, pool_size=100, max_overflow=10, pool_timeout=100, pool_recycle=7200)
        uploadTable.to_sql(updateDBTable, engine, dtype=dtypedict, index=False, if_exists='append')  # append/repalce
    except Exception as e:
        print("上传数据库过程出错！！！", e)


# 获取需要导入数据的月份
def getEveryMonth(nearestMonth, toMonth):
    monthList = []
    nearestMonth = datetime.datetime.strptime(str(nearestMonth)[0:8] + '01', "%Y-%m-%d")  # 月初1号
    while nearestMonth <= toMonth:
        dateStrZZ = nearestMonth.strftime("%Y-%m-%d")
        monthList.append(dateStrZZ)
        nearestMonth += relativedelta(months=1)
    return monthList


# 获取【sfa_bf_fact库】最新日期
connHANA = pymysql.connect(
    host='localhost',
    port=3306,
    user='root',
    passwd='123456',
    db='sfa',
    charset='utf8'
)
cursorHANA = connHANA.cursor()
executeHANA = "SELECT MAX((SY_BFTIME_QD)) FROM sfa_bf_fact WHERE SY_BFTIME_QD != 'nan'"
cursorHANA.execute(executeHANA)
timeStr = cursorHANA.fetchall()
nearestTime = str(timeStr)[3:13].replace(",", "-").replace(" ", "")
connHANA.commit()  # 提交确认
cursorHANA.close()  # 关闭光标
connHANA.close()  # 关闭连接

nearestTimeHANA = datetime.datetime.strptime(nearestTime, "%Y-%m-%d")  # 最新日期
todayTimeHANA = datetime.datetime.now()  # 本日
missMonthHANA = getEveryMonth(nearestTimeHANA, todayTimeHANA)
cprint("【sfa库】最新日期: " + str(nearestTime) + "; 即将导入" + str(missMonthHANA) + "时间段内的数据", 'magenta', attrs=['bold', 'reverse', 'blink'])
for iMonth in missMonthHANA:
    Year = iMonth[0:4]  # 年
    Month = iMonth[5:7]  # 月

    try:
        print(" > 获取HANA数据中")
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
                'SELECT * FROM "HD-HAND.SD.POWER_BI::CV_ZHRWQBF11" WHERE YEAR("CREATEDTIME")=:1 AND MONTH("CREATEDTIME")=:2', [Year, Month])
            matBF = cursorBF.fetchall()
            return matBF

        # 获取签退表指定时间段数据
        def get_matQT(connQT):
            cursorQT = connQT.cursor()
            cursorQT.execute(
                'SELECT * FROM "HD-HAND.SD.POWER_BI::CV_ZHRWQBF12" WHERE YEAR("CREATEDTIME")=:1 AND MONTH("CREATEDTIME")=:2', [Year, Month])
            matQT = cursorQT.fetchall()
            return matQT

        conn = get_HANA_Connection()
        dataBF = pd.DataFrame(get_matBF(conn))
        dataQT = pd.DataFrame(get_matQT(conn))

        endGetHanaDataTime = datetime.datetime.now()
        cprint(Year + "." + Month + "数据成功获取, 拜访表数据有" + str(len(dataBF)) + "行; 签退表数据有" + str(len(dataQT)) + "行; 耗时：" + strftime("%H:%M:%S",
                                                                                                                               gmtime((endGetHanaDataTime - startGetHanaDataTime).seconds)), 'magenta', attrs=['bold', 'reverse', 'blink'])
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

        dataBF['SY_BFTIME_QD'] = pd.to_datetime(dataBF['SY_BFTIME_QD'])
        dataQT['SY_QTTIME_QT'] = pd.to_datetime(dataQT['SY_QTTIME_QT'])

        '''
        数据预处理（清洗掉拜访时间、签退时间、人员编号和门店编号异常的数据）
        '''
        # delVarListBF = ['SY_SFA_ID_QD', 'STAFF_ID_QD', 'SY_BFTIME_QD', 'SY_TYPE_QD']
        # for delvarBF in delVarListBF:
        #     for i in range(len(dataBF[delvarBF])):
        #         if dataBF.loc[i, delvarBF] == 0 or dataBF.loc[i, delvarBF] == '':
        #             dataBF.loc[i, delvarBF] = np.nan
        # dataBF = dataBF.dropna(axis=0, subset=delVarListBF)
        #
        # delVarListQT = ['SY_SFA_ID_QT', 'STAFF_ID_QT', 'SY_QTTIME_QT']
        # for delvarQT in delVarListQT:
        #     for j in range(len(dataQT[delvarQT])):
        #         if dataQT.loc[j, delvarQT] == 0 or dataQT.loc[j, delvarQT] == '':
        #             dataQT.loc[j, delvarQT] = np.nan
        # dataQT = dataQT.dropna(axis=0, subset=delVarListQT)

        # dataBF.to_excel('D:/JR&Zeus_Project/zeus/mergeTable_BF_QT/BF/BF_bei.xlsx', index=False)
        # dataQT.to_excel('D:/JR&Zeus_Project/zeus/mergeTable_BF_QT/QT/QT_bei.xlsx', index=False)
        #
        # df1 = pd.read_excel('D:/JR&Zeus_Project/zeus/mergeTable_BF_QT/BF/BF_bei.xlsx', sheet_name=0, header=0,
        #                     index_col=None)
        # df2 = pd.read_excel('D:/JR&Zeus_Project/zeus/mergeTable_BF_QT/QT/QT_bei.xlsx', sheet_name=0, header=0,
        #                     index_col=None)

        df1 = dataBF.reset_index(drop=True)
        df2 = dataQT.reset_index(drop=True)

        df1.insert(0, 'indexQD', '')
        for i in range(len(df1['SY_BFTIME_QD'])):
            df1.loc[i, "indexQD"] = str(df1.loc[i, 'SY_BFTIME_QD'])[0:11] + str(df1.loc[i, 'STAFF_ID_QD']) + str(
                df1.loc[i, 'SY_SFA_ID_QD'])[0:9]

        df2.insert(0, 'indexQT', '')
        for j in range(len(df2['SY_QTTIME_QT'])):
            df2.loc[j, 'indexQT'] = str(df2.loc[j, 'SY_QTTIME_QT'])[0:11] + str(df2.loc[j, 'STAFF_ID_QT']) + str(
                df2.loc[j, 'SY_SFA_ID_QT'])[0:9]

        '''
        构造拜访表与签退表唯一的主键
        '''
        # 拜访表唯一主键构造
        startBFTime = datetime.datetime.now()
        print(" > 拜访表主键唯一化")
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
        print(">> 拜访表主键唯一化成功, 耗时: " + strftime("%H:%M:%S", gmtime((endBFTime - startBFTime).seconds)))

        # 签退表唯一主键
        startQTTime = datetime.datetime.now()
        print(" > 签退表主键唯一化")
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
        print(">> 签退表主键唯一化成功, 耗时: " + strftime("%H:%M:%S", gmtime((endQTTime - startQTTime).seconds)))

        result = pd.merge(df1.drop_duplicates(), df2.drop_duplicates(), how='left', left_on='UNIQUE_KEYS', right_on='UNIQUE_KEYS')
        # 多级排序，ascending=False代表按降序排序，na_position='last'代表空值放在最后一位
        result.sort_values(by=['UNIQUE_KEYS', 'SY_QTTIME_QT'], ascending=False, na_position='last')
        result.drop_duplicates(subset='UNIQUE_KEYS', keep='last', inplace=True)
        res = result.drop(columns=['indexQD', 'indexQT', 'markQD'])
        # res.to_excel('D:/JR&Zeus_Project/zeus/mergeTable_BF_QT/merge_BF_Table.xlsx', index=False)
        mergeData = res.reset_index(drop=True)
        endMergeTableTime = datetime.datetime.now()
        cprint(Year + "." + Month + "拜访签退表合并成功, 匹配成功" + str(len(res)) + "行; 耗时: " + strftime("%H:%M:%S", gmtime(
            (endMergeTableTime - startMergeTableTime).seconds)), 'magenta', attrs=['bold', 'reverse', 'blink'])
    except Exception as e:
        print("数据处理or表格合并出错！！！", e)

    try:
        startUploadDBTime = datetime.datetime.now()
        print(" > 上传数据库中,请稍等片刻")

        # def getExcel(location):
        #     allExcel = glob.glob(location + "*.xlsx")  # 定位所有xls文件
        #     print("该目录下有" + str(len(allExcel)) + "个excel文件：")
        #     if (len(allExcel) == 0):
        #         return 0
        #     else:
        #         for i in range(len(allExcel)):
        #             print(allExcel[i])
        #         return allExcel

        def newDF():
            createDF = pd.DataFrame(
                columns=['UNIQUE_KEYS', 'MANDT_QD', 'OBJECTI_QD', 'NAME_QD', 'STAFF_DESC_QD', 'SY_DEPT_DESC_QD',
                         'SY_SFA_DESC_QD', 'SY_POINT_DESC_QD', 'F0000003_QD', 'SY_SFA_BINARYCODE_QD', 'SY_TYPE_QD',
                         'SY_SFA_ID_QD', 'SY_BFDATE_QD', 'key_QD', 'CREATEDBY_QD', 'SY_BFTIME_QD', 'MODIFIEDBY_QD',
                         'MODIFIEDTIME_QD', 'CREATEDBYOBJECT_QD', 'OWNERIDOBJECT_QD', 'OWNERDEPTIDOBJECT_QD',
                         'MODIFIEDBYOBJECT_QD', 'SY_LATITUDE_QD', 'SY_LONGITUDE_QD', 'STAFF_ID_QD', 'SY_DEPT_ID_QD',
                         'MANDT_QT', 'OBJECTID_QT', 'NAME_QT', 'STAFF_DESC_SECOND_QT', 'SY_DEPT_DESC_QT',
                         'SY_POINT_DESC_QT', 'SY_PHOTO_QT', 'SY_SFA_BINARYCODE_QT', 'SY_SFA_DESC_QT', 'SY_QTDATE_QT',
                         'XIAOJIE_QT', 'STAFF_DESC_FIRST_QT', 'SY_QTTIME_QT', 'MODIFIEDBY_QT', 'MODIFIEDTIME_QT',
                         'CREATEDBYOBJECT_QT', 'OWNERIDOBJECT_QT', 'OWNERDEPTIDOBJECT_QT', 'MODIFIEDBYOBJECT_QT',
                         'SY_LATITUDE_QT', 'SY_LONGITUDE_QT', 'STAFF_ID_QT', 'SY_DEPT_ID_QT', 'SY_SFA_ID_QT'])
            return createDF

        # 删除本次上传存在的历史数据
        print(" > 删除" + Year + "." + Month + "历史数据")
        connDel = pymysql.connect(
            host='localhost',
            port=3306,
            user='root',
            passwd='123456',
            db='sfa',
            charset='utf8'
        )
        cursorDel = connDel.cursor()
        executeDel = "DELETE FROM sfa_bf_fact WHERE YEAR(SY_BFDATE_QD)='" + Year + "' AND MONTH(SY_BFDATE_QD)='" + Month + "'"
        cursorDel.execute(executeDel)
        delRowNum = cursorDel.rowcount
        connDel.commit()  # 提交确认
        cursorDel.close()  # 关闭光标
        connDel.close()  # 关闭连接
        print(">> 删除" + Year + "." + Month + "时间段内旧数据" + str(delRowNum) + "行")
        # locPathExcel = getExcel("D:\\JR&Zeus_Project\\zeus\\mergeTable_BF_QT\\")  # 获取路径下全部xlsx文件
        uploadData = newDF()
        # for i in range(len(locPathExcel)):
        #     print("加载" + str(locPathExcel[i]) + "数据ing......")
        #     mergeData = pd.read_excel(locPathExcel[i], sheet_name=0, dtype=str)
        #     uploadData = uploadData.append([mergeData])  # 所有数据表合并在uploadData表中
        # mergeData = pd.read_excel('C:/Users/Long/Desktop/BF.xlsx', sheet_name=0, header=0)
        uploadData = uploadData.append([mergeData])
        # okGo = uploadData.astype(str)  # 数据改为字符串类型
        okGo = uploadData.astype(str)
        print(">> 数据上传MySQLing:root -> sfa -> sfa_bf_fact")
        uploadSQL(okGo, 'sfa', 'sfa_bf_fact')  # 上传数据
        endUploadDBTime = datetime.datetime.now()
        cprint("【sfa库】" + Year + "." + Month + "删除旧数据" + str(delRowNum) + "行; 导入新数据" + str(
            len(uploadData)) + "行; 新增数据" + str(len(uploadData) - delRowNum) + "行; 耗时: " + strftime("%H:%M:%S", gmtime(
                (endUploadDBTime - startUploadDBTime).seconds)), 'cyan', attrs=['bold', 'reverse', 'blink'])
    except Exception as e:
        print("上传数据库出错！！！", e)

endTime = datetime.datetime.now()
print(">>>总耗时: " + strftime("%H:%M:%S", gmtime((endTime - startTime).seconds)))
