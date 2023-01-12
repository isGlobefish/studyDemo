# -*- coding: utf-8 -*-
'''
@createTool    : PyCharm-2020.3.2
@projectName   : pythonProjectPy3.9
@originalAuthor: Made in win10.Sys design by deHao.Zou
@createTime    : 2021/1/8 14:47
'''

import pandas as pd
import numpy as np
import xlrd
import xlwt
import glob
import time
import os
import re
import time
import xlwt
import glob
import xlrd
import shutil
import pymysql
import calendar
import datetime
import win32api as ap
import pyautogui as pg
import pyperclip as cp  # 复制粘贴
from termcolor import cprint
import win32com.client as win32
from time import strftime, gmtime

# 打卡记录
origData = pd.read_excel('D:/DataCenter/VisImage/近七天拜访签退.xlsx')
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
        columns=['Name', 'Type', 'sumRecord', 'sumSFA', 'recordDaySub1', 'SFADaySub1', 'daySub7', 'daySub6', 'daySub5', 'daySub4', 'daySub3',
                 'daySub2', 'daySub1', 'partner'])
    return dfNew


# 字典匹配
sub1SFADict = {}  # 前天拜访门店数
sumSFADict = {}  # 总拜访门店数
partnerDict = {}  # 伙伴
sfaRow = len(dealDF)
for i in range(0, sfaRow):
    sub1SFADict[dealDF.loc[i, 'Name']] = dealDF.loc[i, 'daySub1']
    sumSFADict[dealDF.loc[i, 'Name']] = dealDF.loc[i, 'sumSFA']
dictionary = xlrd.open_workbook('D:/DataCenter/VisImage/伙伴对照表.xlsx')
sheet = dictionary.sheet_by_name('Sheet1')
partnerRow = sheet.nrows
for i in range(1, partnerRow):
    values = sheet.row_values(i)
    partnerDict[values[1]] = values[2]

readyTable = newDF()
print(">  归一化数据表")
readyTable['Name'] = origDF['Name']
readyTable['Type'] = origDF['Type']
readyTable['sumRecord'] = origDF['sumRecord']
readyTable['recordDaySub1'] = origDF['daySub1']
readyTable['daySub7'] = origDF['daySub7']
readyTable['daySub6'] = origDF['daySub6']
readyTable['daySub5'] = origDF['daySub5']
readyTable['daySub4'] = origDF['daySub4']
readyTable['daySub3'] = origDF['daySub3']
readyTable['daySub2'] = origDF['daySub2']
readyTable['daySub1'] = origDF['daySub1']
readyTable['partner'] = readyTable.apply(lambda x: partnerDict.setdefault(x['Name'], 0), axis=1)
readyTable['SFADaySub1'] = readyTable.apply(lambda x: sub1SFADict.setdefault(x['Name'], 0), axis=1)
readyTable['sumSFA'] = readyTable.apply(lambda x: sumSFADict.setdefault(x['Name'], 0), axis=1)

# 多级排序，ascending=False代表按降序排序，na_position='last'代表空值放在最后一位
readyExport = readyTable.sort_values(by=['sumRecord', 'sumSFA'], ascending=False, na_position='last')

# readyExport.to_excel('C:/Users/Zeus/Desktop/321.xlsx', index=False)
partnerTable = readyExport[readyExport['partner'] == '邵远登'].reset_index(drop=True)

sub7Time = (datetime.datetime.now() + datetime.timedelta(days=-7)).strftime("%Y-%m-%d")  # 前七天
sub6Time = (datetime.datetime.now() + datetime.timedelta(days=-6)).strftime("%Y-%m-%d")  # 前七天
sub5Time = (datetime.datetime.now() + datetime.timedelta(days=-5)).strftime("%Y-%m-%d")  # 前七天
sub4Time = (datetime.datetime.now() + datetime.timedelta(days=-4)).strftime("%Y-%m-%d")  # 前七天
sub3Time = (datetime.datetime.now() + datetime.timedelta(days=-3)).strftime("%Y-%m-%d")  # 前七天
sub2Time = (datetime.datetime.now() + datetime.timedelta(days=-2)).strftime("%Y-%m-%d")  # 前七天
sub1Time = (datetime.datetime.now() + datetime.timedelta(days=-1)).strftime("%Y-%m-%d")  # 前一天


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


# 创建表格模板并写入数据
f = xlwt.Workbook()  # 创建工作簿
sheet1 = f.add_sheet('Sheet1', cell_overwrite_ok=True)

# 设置列宽
for i in range(13):
    if i == 2 or i == 3:
        sheet1.col(i).width = 256 * 18
    else:
        sheet1.col(i).width = 256 * 14

# 设置行高
for rowi in range(0, 3):
    rowHeight = xlwt.easyxf('font:height 220;')  # 36pt,类型小初的字号
    rowNum = sheet1.row(rowi)
    rowNum.set_style(rowHeight)

# 第A列
sheet1.write_merge(0, 2, 0, 0, "姓名",
                   set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
# 第B列
sheet1.write_merge(0, 2, 1, 1, "打卡类型",
                   set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
# 第C-D列
sheet1.write_merge(0, 0, 2, 3, "2021-01-06至2021-01-12",
                   set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
# 第C列
sheet1.write_merge(1, 2, 2, 2, "打卡记录总次数/条",
                   set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
# 第D列
sheet1.write_merge(1, 2, 3, 3, "拜访门店总数量/家",
                   set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
# 第E列
sheet1.write_merge(0, 0, 4, 5, "2021-01-12",
                   set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
sheet1.write_merge(1, 2, 4, 4, "打卡记录/条",
                   set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
# 第F列
sheet1.write_merge(1, 2, 5, 5, "拜访门店数/家",
                   set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
# 第G-M列
sheet1.write_merge(0, 1, 6, 12, "近七天拜访打卡记录",
                   set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
# 第G列
sheet1.write_merge(2, 2, 6, 6, "2021-01-06",
                   set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
# 第H列
sheet1.write_merge(2, 2, 7, 7, "2021-01-07",
                   set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
# 第I列
sheet1.write_merge(2, 2, 8, 8, "2021-01-08",
                   set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
# 第J列
sheet1.write_merge(2, 2, 9, 9, "2021-01-09",
                   set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
# 第K列
sheet1.write_merge(2, 2, 10, 10, "2021-01-10",
                   set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
# 第L列
sheet1.write_merge(2, 2, 11, 11, "2021-01-11",
                   set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
# 第M列
sheet1.write_merge(2, 2, 12, 12, "2021-01-12",
                   set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))

rowLen = len(partnerTable)
sendcColumns = ['Name', 'Type', 'sumRecord', 'sumSFA', 'recordDaySub1', 'SFADaySub1', 'daySub7', 'daySub6', 'daySub5', 'daySub4', 'daySub3',
                'daySub2', 'daySub1']
# 填充表格数据
for rowi in range(rowLen):
    for colj, colName in enumerate(sendcColumns):
        sheet1.write_merge(rowi + 3, rowi + 3, colj, colj, str(partnerTable.loc[rowi, colName]),
                           set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=1))

f.save('C:/Users/Zeus/Desktop/表头设计.xls')  # 保存文件


# Excel格式转换：.xls ---> .xlsx
def convFormat(openPath, savePath):
    fileList = os.listdir(openPath)  # 该文件夹下所有的文件（包括文件夹）
    print("转换" + str(fileList) + "文件格式")
    for file in fileList:  # 遍历所有文件
        fileName = os.path.splitext(file)[0]  # 获取文件名
        fileType = os.path.splitext(file)[1]  # 获取文件扩展名
        openFiles = openPath + fileName + fileType
        saveFiles = savePath + fileName + fileType
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        wb = excel.Workbooks.Open(openFiles)
        wb.SaveAs(saveFiles + "x", FileFormat=51)  # FileFormat = 51转为.xlsx、FileFormat = 56转为.xls
        wb.Close()
        excel.Application.Quit()
    print("（xls->xlsx）已转换格式完成！！！")


convFormat('C:\\Users\\Zeus\\Desktop\\999\\', 'C:\\Users\\Zeus\\Desktop\\999\\')  # 转换表格格式(.xls -> .xlsx)

from win32com.client import Dispatch, DispatchEx
import pythoncom
from PIL import ImageGrab, Image
import uuid


# screenArea——格式类似"A1:J10"
def excelCatchScreen(file_name, sheet_name, screen_area, save_path, img_name=False):
    pythoncom.CoInitialize()  # excel多线程相关
    excel = DispatchEx("Excel.Application")  # 启动excel
    excel.Visible = True  # 可视化
    excel.DisplayAlerts = False  # 是否显示警告
    wb = excel.Workbooks.Open(file_name)  # 打开excel
    ws = wb.Sheets(sheet_name)  # 选择Sheet
    ws.Range(screen_area).CopyPicture()  # 复制图片区域
    ws.Paste()  # 粘贴 ws.Paste(ws.Range('B1'))  # 将图片移动到具体位置

    # name = str(uuid.uuid4())  # 重命名唯一值
    name = "拜访数据可视化"
    new_shape_name = name[:6]
    excel.Selection.ShapeRange.Name = new_shape_name  # 将刚刚选择的Shape重命名, 避免与已有图片混淆

    ws.Shapes(new_shape_name).Copy()  # 选择图片
    img = ImageGrab.grabclipboard()  # 获取剪贴板的图片数据
    if not img_name:
        img_name = name + ".PNG"
    img.save(save_path + img_name)  # 保存图片
    wb.Close(SaveChanges=0)  # 关闭工作薄，不保存
    excel.Quit()  # 退出excel
    pythoncom.CoUninitialize()


if __name__ == '__main__':
    excelCatchScreen("C:/Users/Zeus/Desktop/999/表头设计.xlsx", "Sheet1", "A1:M16", 'D:/DataCenter/VisImage/')
