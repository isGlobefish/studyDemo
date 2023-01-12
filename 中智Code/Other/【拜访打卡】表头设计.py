# -*- coding: utf-8 -*-
'''
@createTool    : PyCharm-2020.3.2
@projectName   : pythonProject
@originalAuthor: Made in win10.Sys design by deHao.Zou
@createTime    : 2021/01/26 11:25


from win32com.client import Dispatch, DispatchEx
import pythoncom
from PIL import ImageGrab, Image
import uuid


# screenArea——格式类似"A1:J10"
def excelCatchScreen(file_name, sheet_name, screen_area, save_path, img_name=False):
    pythoncom.CoInitialize()  # excel多线程相关
    excel = DispatchEx("Excel.Application")  # 启动excel
    excel.Visible = False  # 可视化
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
    excelCatchScreen("C:/Users/Zeus/Desktop/321.xlsx", "Sheet1", "A1:M16", 'D:/DataCenter/VisImage/')
'''
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
def write_excel():
    try:
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

        # 读取下载数据
        import pandas as pd
        saveData = pd.read_excel('C:/Users/Long/Desktop/ZFI77/downloadFiles/export.XLSX', header=0)
        rowLen = len(saveData)

        # 填充表格数据
        for rowj in range(rowLen):
            sheet1.write_merge(rowj + 5, rowj + 5, 2, 2, '%.2f' % saveData.loc[rowj, "本期金额"],
                               set_style('楷体_GB2312', 200, False, Halign=0, Valign=0, setBorder=1, setbgcolor=1))
            sheet1.write_merge(rowj + 5, rowj + 5, 3, 3, '%.2f' % saveData.loc[rowj, "本年累计金额"],
                               set_style('楷体_GB2312', 200, False, Halign=0, Valign=0, setBorder=1, setbgcolor=1))

        f.save('C:/Users/Zeus/Desktop/表头设计.xls')  # 保存文件

    except Exception as e:
        print("构造表或写入数据出错！！！", e)


if __name__ == '__main__':
    write_excel()
