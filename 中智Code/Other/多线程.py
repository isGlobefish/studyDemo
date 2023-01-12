# -*- coding: utf-8 -*-
'''
@createTool    : PyCharm-2020.2.2
@projectName   : pythonProjectPy3.9 
@originalAuthor: Made in win10.Sys design by deHao.Zou
@createTime    : 2020/12/7 15:38
'''

# import pandas as pd
#
# data = pd.read_excel('C:/Users/Long/Desktop/DKH对照表.xlsx', sheet_name=1)
#
# dataSplit = pd.DataFrame((x.split('-') for x in data['OUT_Dept_1']), index=data.index, columns=['leftSplit', 'rightSplit'])
# dataSplit.head(5)
#
# dataMerge = pd.merge(data, dataSplit, right_index=True, left_index=True)
# dataMerge.head(5)

import glob
import os
import pandas as pd
import calendar
from datetime import datetime


def newdf():
    df = pd.DataFrame(
        columns=['desc_1', 'DKH_materiel_id', 'materiel_desc', 'norms', 'date', 'sfa_desc', 'sfa_id', 'amount',
                 'UnitPrice', 'sales_Money', 'bz_UnitPrice', 'bz_sales_Money', 'customer', 'materiel_alias', 'dept_5',
                 'sfa_client_desc', 'state', 'client_id', 'client_desc', 'client_alias'])
    return df


Year = input("输入数据删除的年份：")
Month = input("输入数据删除的月份：").zfill(2)
firstDay = '01'  # 前一个月第一天
lastDay = str(calendar.monthrange(int(Year), int(Month))[1]).zfill(2)  # 前一月最后一日


# 指定月份数据上传
def dkhOrginalFiles(path):
    dkhFileList = os.listdir(path)
    all_Xls = glob.glob(path + "*.xls")
    all_Xlsx = glob.glob(path + "*.xlsx")
    all_Csv = glob.glob(path + "*.csv")
    print("该目录下有" + '\n' + str(dkhFileList) + ";" + '\n' + "其中【xls:" + str(len(all_Xls)) + ", xlsx:" + str(
        len(all_Xlsx)) + ", csv:" + str(len(all_Csv)) + "】")
    dfCreate = newdf()
    for dkhindex, dkhfile in enumerate(dkhFileList, start=1):  # 遍历所有文件
        fileName = os.path.splitext(dkhfile)[0]  # 获取文件名
        fileType = os.path.splitext(dkhfile)[1]  # 获取文件扩展名
        fileFullPath = path + fileName + fileType  # 文件完整路径
        print(str(dkhindex).zfill(2) + '/' + str(len(dkhFileList)) + " 数据读取进程：" + dkhfile)
        togetData = pd.read_excel(fileFullPath, header=0)
        togetData['date'] = pd.to_datetime(togetData['date'])
        selectData = togetData[(togetData['date'] >= pd.to_datetime(Year + Month + firstDay)) & (
                togetData['date'] <= pd.to_datetime(Year + Month + lastDay))]
        selectData = selectData.reset_index(drop=True)
        dfCreate = dfCreate.append([selectData])  # 合并数据
    dfCreate.to_excel("C:/Users/Long/Desktop/123.xlsx", index=False)


dkhOrginalFiles('G:\\大客户数据源\\')
