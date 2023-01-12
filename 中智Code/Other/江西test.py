# -*- coding: utf-8 -*-
'''
@createTool    : PyCharm-2020.2.2
@projectName   : pythonProjectPy3.9 
@originalAuthor: Made in win10.Sys design by deHao.Zou
@createTime    : 2020/12/23 14:54
'''

import time
import datetime
import calendar
import pymysql
import pandas as pd
from sqlalchemy import create_engine  # 连接mysql使用
from sqlalchemy.types import NVARCHAR
from termcolor import cprint
from time import strftime, gmtime
from dateutil.relativedelta import relativedelta
import time
import hmac
import json
import base64
import hashlib
import requests
import calendar
import urllib.parse
import urllib.request
from datetime import datetime

today = int(time.strftime("%d", time.localtime()))
releaseTime = str(datetime.now()).split('.')[0]  # 发布时间
if today <= 2:
    Year = str(time.strftime("%Y", time.localtime())).zfill(4)  # 本年
    Month = str(int(time.strftime("%m", time.localtime())) - 1).zfill(2)  # 前一月
    if today == 1:
        daySub1 = str(calendar.monthrange(int(Year), int(Month))[1]).zfill(2)  # 前一月最后一日
        daySub2 = str(int(calendar.monthrange(int(Year), int(Month))[1]) - 1).zfill(2)  # 前一月倒数第二天
    elif today == 2:
        daySub1 = str(int(time.strftime("%d", time.localtime())) - 1).zfill(2)  # 前一日
        daySub2 = str(calendar.monthrange(int(Year), int(Month))[1]).zfill(2)  # 前一月最后一日
else:
    Year = str(time.strftime("%Y", time.localtime())).zfill(4)  # 本年
    Month = str(int(time.strftime("%m", time.localtime()))).zfill(2)  # 本月
    daySub1 = str(int(time.strftime("%d", time.localtime())) - 1).zfill(2)  # 前一日
    daySub2 = str(int(time.strftime("%d", time.localtime())) - 2).zfill(2)  # 前两日

conn = pymysql.connect(host='192.168.249.150',  # 数据库地址
                       port=3306,  # 数据库端口
                       user='alex',  # 用户名
                       passwd='123456',  # 数据库密码
                       db='dkh',  # 数据库名
                       charset='utf8')  # 字符串类型
cursor = conn.cursor()
executeCode = """SELECT * FROM dkhfact where customer = '益丰' AND YEAR(date) = '""" + Year + """' AND MONTH(date) = '""" + Month + """' AND (desc_1 = '江西' OR desc_1 = '江西天顺') UNION ALL
SELECT * FROM dkhfact where customer = '大参林' AND YEAR(date) = '""" + Year + """' AND MONTH(date) = '""" + Month + """' AND (desc_1 = '赣州' OR desc_1 = '南昌') UNION ALL
SELECT * FROM dkhfact where customer = '高济' AND YEAR(date) = '""" + Year + """' AND MONTH(date) = '""" + Month + """' AND (desc_1 = '江西开心人大药房连锁有限公司');"""
cursor.execute(executeCode)  # 执行查询
rowNum = cursor.rowcount  # 查询数据条数
Data = cursor.fetchall()  # 获取全部查询数据
conn.commit()  # 提交确认
cursor.close()  # 关闭光标
conn.close()  # 关闭连接

colNames = ['省/市/公司名', '商品编码', '商品名称', '商品规格', '销售日期', '门店名称', '门店编码', '数量', '单价',
            '零售金额', '标准单价', '标准零售金额', '客户体系', '商品简称', '流向编码/名称', '门店客户名称',
            '数据状态', '客户编码', '客户名称', '客户简称']
jiangxiData = pd.DataFrame(Data, columns=colNames)
# jiangxiData['销售日期'] = pd.to_datetime(jiangxiData['销售日期'], format='%Y/%m/%d')
jiangxiData['销售日期'] = jiangxiData['销售日期'].dt.date
jiangxiData.to_excel('C:/Users/Long/Desktop/江西202012.xlsx', index=False)
