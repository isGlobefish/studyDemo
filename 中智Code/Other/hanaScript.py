# -*- coding: utf-8 -*-
'''
@createTool    : PyCharm-2020.2.2
@projectName   : pythonProject 
@originalAuthor: Made in win10.Sys design by deHao.Zou
@createTime    : 2020/10/22 10:15
'''
import pandas as pd
import numpy as np
import pyhdb  # 加载连接HANA的所需模块
import re
from xlrd import xldate_as_tuple
import openpyxl
import glob
import time
import datetime
import time
import datetime
import calendar
import pymysql
from sqlalchemy import create_engine  # 连接mysql使用
from sqlalchemy.types import NVARCHAR
from termcolor import cprint
from time import strftime, gmtime
from dateutil.relativedelta import relativedelta

Year = input("输入年份:").zfill(4)
Month = input("输入月份:").zfill(2)


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
    host='192.168.249.150',
    port=3306,
    user='alex',
    passwd='123456',
    db='sfa',
    charset='utf8'
)
cursorHANA = connHANA.cursor()
executeHANA = "SELECT MAX(SY_BFTIME_QD) FROM sfa_bf_fact"
cursorHANA.execute(executeHANA)
timeStr = cursorHANA.fetchall()
nearestTime = str(timeStr)[20:32].replace(",", "-").replace(" ", "")
connHANA.commit()  # 提交确认
cursorHANA.close()  # 关闭光标
connHANA.close()  # 关闭连接

nearestTimeHANA = datetime.datetime.strptime(nearestTime, "%Y-%m-%d")  # 最新日期

todayTimeHANA = datetime.datetime.now()  # 本日

missMonthHANA = getEveryMonth(nearestTimeHANA, todayTimeHANA)

for iMonth in missMonthHANA:
    Year = iMonth[0:4]  # 年
    Month = iMonth[5:7]  # 月

# 获取【liaochengfenxi库】最新日期
connHANA = pymysql.connect(
    host='192.168.249.150',
    port=3306,
    user='alex',
    passwd='123456',
    db='liaochengfenxi',
    charset='utf8'
)
cursorHANA = connHANA.cursor()
executeHANA = "SELECT MAX(`日期`) FROM liaocheng_sale_fact WHERE `公司`='中智Code'"
cursorHANA.execute(executeHANA)
timeStr = cursorHANA.fetchall()
nearestTime = str(timeStr)[20:32].replace(",", "-").replace(" ", "")
connHANA.commit()  # 提交确认
cursorHANA.close()  # 关闭光标
connHANA.close()  # 关闭连接

nearestTimeHANA = datetime.datetime.strptime(nearestTime, "%Y-%m-%d")  # 最新日期

todayTimeHANA = datetime.datetime.now()  # 本日

missMonthHANA = getEveryMonth(nearestTimeHANA, todayTimeHANA)

for iMonth in missMonthHANA:
    Year = iMonth[0:4]  # 年
    Month = iMonth[5:7]  # 月

# 获取【pbi_log库】最新日期
connHANA = pymysql.connect(
    host='192.168.249.150',
    port=3306,
    user='alex',
    passwd='123456',
    db='powerbi',
    charset='utf8'
)
cursorHANA = connHANA.cursor()
executeHANA = "SELECT MAX(Date) FROM pbi_log"
cursorHANA.execute(executeHANA)
timeStr = cursorHANA.fetchall()[0]
nearestTime = str(timeStr)[2:13].replace(",", "-").replace(" ", "")
connHANA.commit()  # 提交确认
cursorHANA.close()  # 关闭光标
connHANA.close()  # 关闭连接

nearestTimeHANA = datetime.datetime.strptime(nearestTime, "%Y-%m-%d")  # 最新日期

todayTimeHANA = datetime.datetime.now()  # 本日

missMonthHANA = getEveryMonth(nearestTimeHANA, todayTimeHANA)

for iMonth in missMonthHANA:
    Year = iMonth[0:4]  # 年
    Month = iMonth[5:7]  # 月


# 获取 Connection 对象
def get_HANA_Connection():
    connObj = pyhdb.connect(
        host="192.168.20.197",  # HANA地址
        port=30015,  # HANA端口号
        user="HANA110731",  # 用户名
        password="Zeus_110731"  # 密码
    )
    return connObj


# 获取拜访表指定时间段数据
def get_matBF(conn):
    cursorBF = conn.cursor()
    cursorBF.execute(
        'SELECT * FROM "HD-HAND.SD.POWER_BI::CV_ZHRWQBF11" WHERE YEAR("DATE_SQL")=:1 AND MONTH("DATE_SQL")=:2',
        [Year, Month])
    matBF = cursorBF.fetchall()
    return matBF


# 获取签退表指定时间段数据
def get_matQT(conn):
    cursorQT = conn.cursor()
    cursorQT.execute(
        'SELECT * FROM "HD-HAND.SD.POWER_BI::CV_ZHRWQBF12" WHERE YEAR("DATE_SQL")=:1 AND MONTH("DATE_SQL")=:2',
        [Year, Month])
    matQT = cursorQT.fetchall()
    return matQT


conn = get_HANA_Connection()
dataBF = pd.DataFrame(get_matBF(conn))

# 删除本次上传存在的历史数据
connDel = pymysql.connect(
    host='192.168.249.150',
    port=3306,
    user='alex',
    passwd='123456',
    db='sfa',
    charset='utf8'
)
cursorDel = connDel.cursor()
executeDel = "SELECT * FROM sfa_bf_fact WHERE YEAR(SY_BFTIME_QD)='" + Year + "' AND	MONTH(SY_BFTIME_QD)='" + Month + "'"
cursorDel.execute(executeDel)
delRowNumHana = cursorDel.rowcount
dataDel = cursorDel.fetchall()
connDel.commit()  # 提交确认
cursorDel.close()  # 关闭光标
connDel.close()  # 关闭连接

data = pd.read_excel('C:/Users/Long/Desktop/JT2020.xlsx', header=0)
res = data.drop(columns=['DKH_materiel_id', 'materiel_desc', 'norms'])

res2 = res.copy()

# # 按月列表
# import datetime
#
# firstDate = datetime.datetime.strptime('2020-10-23', "%Y-%m-%d")
#
# from dateutil.relativedelta import relativedelta
#
# print((firstDate + relativedelta(months=1)).strftime("%Y-%m"))
#
# # 全亿 and 中智Code
# import time
# import datetime
# import calendar
# import pymysql
# import pandas as pd
# from sqlalchemy import create_engine  # 连接mysql使用
# from sqlalchemy.types import NVARCHAR
# from termcolor import cprint
# from time import strftime, gmtime
# from dateutil.relativedelta import relativedelta
#
# Z0Z = pymysql.connect(
#     host='192.168.249.150',
#     port=3306,
#     user='alex',
#     passwd='123456',
#     db='liaochengfenxi',
#     charset='utf8'
# )
# cursorZZ = Z0Z.cursor()
# toYear = datetime.datetime.now().strftime("%Y")  # 本年
# executeZZ = "SELECT `日期` FROM liaocheng_sale_fact WHERE `公司`='中智Code'AND YEAR(`日期`)='" + toYear + "'"  # 年初注意要修改
# cursorZZ.execute(executeZZ)
# nearestTimeZZ = max(cursorZZ.fetchall())[0]  # 最晚日期
# Z0Z.commit()  # 提交确认
# cursorZZ.close()  # 关闭光标
# Z0Z.close()  # 关闭连接
#
# todayTime = datetime.datetime.now()  # 本日
#
#
# def getEveryMonth(nearestMonthZZ, toMonthZZ):
#     monthListZZ = []
#     nearestMonthZZ = datetime.datetime.strptime(str(nearestMonthZZ)[0:8] + '01', "%Y-%m-%d")
#     while nearestMonthZZ <= toMonthZZ:
#         dateStrZZ = nearestMonthZZ.strftime("%Y-%m-%d")
#         monthListZZ.append(dateStrZZ)
#         nearestMonthZZ += relativedelta(months=1)
#     return monthListZZ
#
#
# missMonth = getEveryMonth(nearestTimeZZ, todayTime)
# for iDate in missMonth:
#     Year = iDate[0:4]  # 年
#     Month = iDate[5:7]  # 月
#     Day = iDate[8:10]  # 日
#
# firstDate = datetime.datetime.strptime('2020-12-01', "%Y-%m-%d")
# firstDate1 = str(datetime.datetime.strptime('2020-10-25', "%Y-%m-%d"))
# getEveryMonth(firstDate1, firstDate)
#
#
# def getEveryDay(beginDate, endDate):
#     dateList = []
#     firstDate = datetime.datetime.strptime(beginDate, "%Y-%m-%d")
#     lastDate = datetime.datetime.strptime(endDate, "%Y-%m-%d")
#     while firstDate <= lastDate:
#         dateStr = firstDate.strftime("%Y-%m-%d")
#         dateList.append(dateStr)
#         firstDate += datetime.timedelta(days=1)
#     return dateList
#
#
# missDate = getEveryDay(nearestTimeZZ.strip(), todayTime.strip())
# for iDate in missDate:
#     Year = iDate[0:4]
#     Month = iDate[5:7]
#     Day = iDate[8:10]


# # 用户日志
# import time
# import datetime
# import calendar
# import pymysql
# import pandas as pd
# from sqlalchemy import create_engine  # 连接mysql使用
# from sqlalchemy.types import NVARCHAR
# from termcolor import cprint
# from time import strftime, gmtime
#
# Pbi = pymysql.connect(
#     host='192.168.249.150',
#     port=3306,
#     user='alex',
#     passwd='123456',
#     db='powerbi',
#     charset='utf8'
# )
# cursorPbi = Pbi.cursor()
# executePbi = "SELECT Date FROM pbi_log"
# cursorPbi.execute(executePbi)
# nearestTime = max(cursorPbi.fetchall())[0]  # 最晚日期
# Pbi.commit()  # 提交确认
# cursorPbi.close()  # 关闭光标
# Pbi.close()  # 关闭连接
#
# today = datetime.datetime.strptime(datetime.datetime.now().strftime("%Y-%m-%d"), "%Y-%m-%d")  # 本日
# todayBefore = (today + datetime.timedelta(days=-1)).strftime("%Y-%m-%d")  # 前一天
#
# def getEveryDay(beginDate, endDate):
#     dateList = []
#     firstDate = datetime.datetime.strptime(beginDate, "%Y-%m-%d")
#     lastDate = datetime.datetime.strptime(endDate, "%Y-%m-%d")
#     while firstDate <= lastDate:
#         dateStr = firstDate.strftime("%Y-%m-%d")
#         dateList.append(dateStr)
#         firstDate += datetime.timedelta(days=1)
#     return dateList
#
# missDate = getEveryDay(nearestTime.strip(), todayBefore.strip())
# for iDate in missDate:
#     Year = iDate[0:4]
#     Month = iDate[5:7]
#     Day = iDate[8:10]


import xlrd
import xlwt
import os


# Excel格式转换：.csv ---> .xls
def getFormat(openPath, savePath):
    fileList = os.listdir(openPath)  # 该文件夹下所有的文件（包括文件夹）
    print("转换" + str(fileList) + "文件格式（csv->xls）")
    for file in fileList:  # 遍历所有文件
        fileName = os.path.splitext(file)[0]  # 获取文件名
        fileType = os.path.splitext(file)[1]  # 获取文件扩展名
        data = xlrd.open_workbook(openPath + fileName + fileType)
        sheet1Data = data.sheet_by_index(0)
        workbook = xlwt.Workbook(encoding='utf-8')
        booksheet = workbook.add_sheet('Sheet1', cell_overwrite_ok=True)
        nrows = sheet1Data.nrows
        cols = sheet1Data.ncols
        for i in range(nrows):
            for j in range(cols):
                booksheet.write(i, j, sheet1Data.cell_value(rowx=i, colx=j))
        workbook.save(savePath + fileName + '.xls')
    print("已转换格式完成！！！")


getFormat('E:/大客户数据/GD/downloadGD/', 'E:/大客户数据/GD/xlsFormatGD/')

# import xlwt
# workbook=xlwt.Workbook(encoding='utf-8')
# booksheet=workbook.add_sheet('Sheet1', cell_overwrite_ok=True)
# DATA=(('学号','姓名','年龄','性别','成绩'),
#    ('1001','A','11','男','12'),
#    ('1002','B','12','女','22'),
#    ('1003','C','13','女','32'),
#    ('1004','D','14','男','52'),
#    )
# for i,row in enumerate(DATA):
#   for j,col in enumerate(row):
#     booksheet.write(i,j,col)
# workbook.save('C:/Users/Long/Desktop/grade.xls')


import pandas as pd
import numpy as np

data = pd.DataFrame(
    {'a': [1, 2, 4, np.nan, 7, 9], 'b': [np.nan, 'b', np.nan, np.nan, 'd', 'e'], 'c': [np.nan, 0, 4, np.nan, np.nan, 5],
     'd': [np.nan, np.nan, np.nan, np.nan, np.nan, np.nan]})

print(data)

data.dropna(axis=0,subset = ["a", "b"])   # 丢弃指定列中有缺失值的行







