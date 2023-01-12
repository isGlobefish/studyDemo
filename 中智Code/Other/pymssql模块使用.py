# -*- coding: utf-8 -*-
'''
@createTool    : PyCharm-2020.3.3
@projectName   : pythonProjectPy3.9
@originalAuthor: Made in win10.Sys design by deHao.Zou
@createTime    : 2021/3/26 10:38
'''
import pymssql
import pandas as pd

conn = pymssql.connect(host='localhost',
                       port = 1433,
                       user='sa',
                       password='123456',
                       database='SY',
                       charset='utf8')

# 查看连接是否成功
cursor = conn.cursor()
sql = 'select * from ddmonth'
cursor.execute(sql)
# 用一个rs变量获取数据
rs = cursor.fetchall()

data = pd.DataFrame(rs)

print(data)