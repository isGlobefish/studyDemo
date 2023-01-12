# -*- coding: utf-8 -*-
'''
@createTool    : PyCharm-2020.2.2
@projectName   : pythonProject
@originalAuthor: Made in win10.Sys design by deHao.Zou
@createTime    : 2020/11/26 11:40

# 鼠标屏幕定位
import  os
import  time
import  pyautogui as pg
try:
    while True:
        print("按下组合键 {Ctrl}+C 结束执行\n")
        sW, sH = pg.size()  #获取屏幕的尺寸（像素）screenWidth，screenHeight
        print("屏幕分辨率：\n"+str(sW)+','+str(sH)+'\n')  #打印屏幕分辨率
        x,y = pg.position()   #获取当前鼠标的坐标（像素）
        print("鼠标坐标:\n" + str(x).rjust(4)+','+str(y).rjust(4)) #打印鼠标坐标值
        time.sleep(1) #等待1秒
        os.system('cls')   #清屏
except KeyboardInterrupt:
    print('\n结束,按任意键退出....') #检测到Ctrl+c组合键结束运行

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

year = str(int(time.strftime("%Y", time.localtime())) - 1)  # 年
month = str(int(time.strftime("%m", time.localtime())) - 1)  # 下个月


# 清空指定文件夹
def deleteOldFiles(path):
    deleteFileList = os.listdir(path)
    all_Xls = glob.glob(path + "*.xls")
    all_Xlsx = glob.glob(path + "*.xlsx")
    all_Csv = glob.glob(path + "*.csv")
    print("该目录下有" + '\n' + str(deleteFileList) + ";" + '\n' + "其中【xls:" + str(len(all_Xls)) + ", xlsx:" + str(
        len(all_Xlsx)) + ", csv:" + str(len(all_Csv)) + "】")
    if (len(all_Xls) != 0 or len(all_Xlsx) != 0 or len(all_Csv) != 0):
        for deletefile in deleteFileList:
            isDeleteFile = os.path.join(path, deletefile)
            if os.path.isfile(isDeleteFile):
                os.remove(isDeleteFile)
        all_DelXls = glob.glob(path + "*.xls")
        all_DelXlsx = glob.glob(path + "*.xlsx")
        all_DelCsv = glob.glob(path + "*.csv")
        if (len(all_DelXls) == 0 and len(all_DelXlsx) == 0 and len(all_DelCsv) == 0):
            print("已清空文件夹！！！")
        else:
            print("存在未删除文件")
    else:
        print("不存在excel文件")


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
    print(year + month + "（xls->xlsx）已转换格式完成！！！")


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


# 模拟鼠标操作
def simulateMouseOperation():
    try:
        downloadPath = 'C:\\Users\\Long\\Desktop\\ZFI77\\downloadFiles\\'
        deleteOldFiles(downloadPath)  # 清空文件
        ap.ShellExecute(0, 'open', 'D:\\SAP\\SAPgui\\saplogon.exe', '', '', 1)  # 打开SAP Logon
        time.sleep(6)
        pg.press('down', presses=2, interval=0.125)  # 按下 2次
        time.sleep(1)
        pg.press('enter')  # 回车
        time.sleep(3)
        user = '106610'  # 用户账号
        password = 'abc123'  # 密码
        cp.copy(user)  # 复制用户账号
        pg.hotkey('ctrl', 'v')  # 粘贴用户账号
        # pg.typewrite(user)
        time.sleep(1)
        pg.press('down')  # 按下
        cp.copy(password)  # 复制密码
        pg.hotkey('ctrl', 'v')  # 粘贴密码
        # pg.typewrite(password)
        pg.press('enter')  # 回车
        time.sleep(4)
        affcode = 'ZFI77'  # 事务编码
        # cp.copy(affcode)  # 复制事务编码
        # pg.hotkey('ctrl', 'v')  # 粘贴事务编码
        pg.typewrite(affcode)
        pg.press('enter')  # 回车
        time.sleep(3)
        compcode = '2200'  # 公司代码
        cp.copy(compcode)  # 复制公司代码
        pg.hotkey('ctrl', 'v')  # 粘贴公司代码
        # pg.typewrite(compcode)
        pg.press('down')  # 按下
        time.sleep(1)
        finversion = '20180101'  # 财务版本
        cp.copy(finversion)  # 复制财务版本
        pg.hotkey('ctrl', 'v')  # 粘贴财务版本
        # pg.typewrite(finversion)
        pg.press('down')  # 按下
        time.sleep(1)
        finYear = '2019'  # 财务年度
        pg.press('backspace', presses=4, interval=0.12)  # 按backspace 4次
        cp.copy(finYear)  # 复制财务年度
        pg.hotkey('ctrl', 'v')  # 粘贴财务年度
        # pg.typewrite(finYear)
        # pg.press('down')  # 按下
        pg.press('enter')  # 回车
        time.sleep(2)
        pg.press('down')  # 按下
        finMonth = '10'  # 财务期间
        pg.press('backspace', presses=2, interval=0.1)  # 按backspace 2次
        cp.copy(finMonth)  # 复制财务期间
        pg.hotkey('ctrl', 'v')  # 粘贴财务期间
        # pg.typewrite(finMonth)
        pg.press('F8')  # 按F8
        time.sleep(3)
        # pg.click(204, 105, clicks=1)  # 点击列表
        # time.sleep(1)
        # pg.click(226, 179, clicks=1)  # 点击导出
        # time.sleep(1)
        # pg.click(459, 199, clicks=1)  # 点击电子表格
        # 获取当前屏幕分辨率
        screenWidth, screenHeight = pg.size()
        pg.click(screenWidth / 2, screenHeight / 2, clicks=1, button='right')
        time.sleep(1)
        pg.press('down', presses=6, interval=0.1)  # 按下 6次
        time.sleep(1)
        pg.press('enter', presses=2, interval=1)  # 回车继续
        time.sleep(1)
        pg.press('tab', presses=6, interval=0.1)
        # pg.click(212, 283, clicks=1)  # 点击桌面
        time.sleep(1)
        pg.press('down', presses=1, interval=1)
        time.sleep(1)
        pg.press('space', presses=1, interval=1)  # 选中桌面
        time.sleep(1)
        pg.press('tab', presses=1, interval=1)  # 选中第一个文件
        # pg.click(359, 216, clicks=1)  # 点击第一个文件
        # time.sleep(2)
        pg.press('down', presses=7, interval=0.1)  # 按下 7次
        time.sleep(2)
        pg.press('enter')  # 回车
        time.sleep(2)
        pg.press('down', presses=1, interval=1)  # 按下
        time.sleep(2)
        pg.press('enter', presses=1, interval=2)  # 回车 1次
        time.sleep(2)
        pg.press('tab', presses=4, interval=0.5)  # tab 4次
        time.sleep(2)
        pg.press('enter', presses=1, interval=2)  # 回车
        # time.sleep(10)
        # pg.press('left', presses=1, interval=0.1)
        # pg.press('enter', presses=1, interval=0.1)
        # pg.click(1118, 52, clicks=1)  # 关闭
        # time.sleep(2)
        # pg.click(305, 410, clicks=1)  # 是
        # time.sleep(2)
        # pg.click(1451, 13, clicks=1)  # 退出
    except Exception as e:
        print("模拟鼠标操作出错，请修改出错的定位参数！！！", e)


# 创建表格模板并写入数据
def write_excel():
    try:
        f = xlwt.Workbook()  # 创建工作簿
        sheet1 = f.add_sheet(year + month, cell_overwrite_ok=True)  # 创建名为：年+月
        # 设置列宽
        col1 = sheet1.col(0)
        col2 = sheet1.col(1)
        col3 = sheet1.col(2)
        col4 = sheet1.col(3)
        col1.width = 256 * 50
        col2.width = 256 * 3
        col3.width = 256 * 18
        col4.width = 256 * 18
        # 设置行高
        tall_style = xlwt.easyxf('font:height 420;')  # 36pt,类型小初的字号
        row5 = sheet1.row(4)
        row5.set_style(tall_style)
        for rowi in range(5, 44):
            rowHeight = xlwt.easyxf('font:height 255;')  # 36pt,类型小初的字号
            rowNum = sheet1.row(rowi)
            rowNum.set_style(rowHeight)
        # 第1行
        sheet1.write_merge(0, 0, 0, 3, "现金流量表",
                           set_style('楷体_GB2312', 320, False, Halign=0, Valign=0, setBorder=1, setbgcolor=1))
        # 第2行
        sheet1.write_merge(1, 1, 0, 0, "单位:玉林市逸仙中药材有限公司",
                           set_style('楷体_GB2312', 230, False, Halign=1, Valign=0, setBorder=1, setbgcolor=1))
        sheet1.write_merge(1, 1, 1, 3, year + "年" + month + "月",
                           set_style('楷体_GB2312', 230, False, Halign=1, Valign=0, setBorder=1, setbgcolor=1))
        # 第3行
        sheet1.write_merge(2, 2, 2, 3, "中智财A表(二)",
                           set_style('楷体_GB2312', 200, False, Halign=2, Valign=0, setBorder=1, setbgcolor=1))
        # 第4行
        sheet1.write_merge(3, 3, 2, 3, "单位：元",
                           set_style('楷体_GB2312', 230, False, Halign=2, Valign=0, setBorder=1, setbgcolor=1))
        # 第5行
        sheet1.write_merge(4, 4, 0, 0, "项目",
                           set_style('楷体_GB2312', 200, False, Halign=0, Valign=2, setBorder=0, setbgcolor=0))
        sheet1.write_merge(4, 4, 1, 1, "",
                           set_style('楷体_GB2312', 200, False, Halign=0, Valign=2, setBorder=0, setbgcolor=0))
        sheet1.write_merge(4, 4, 2, 2, "本月数",
                           set_style('楷体_GB2312', 200, False, Halign=0, Valign=2, setBorder=0, setbgcolor=0))
        sheet1.write_merge(4, 4, 3, 3, "本年累计数",
                           set_style('楷体_GB2312', 200, False, Halign=0, Valign=2, setBorder=0, setbgcolor=0))
        # 第6行
        sheet1.write_merge(5, 5, 0, 0, "一、经营活动产生的现金流量：",
                           set_style('楷体_GB2312', 200, False, Halign=1, Valign=0, setBorder=0, setbgcolor=1))
        sheet1.write_merge(5, 5, 1, 1, "",
                           set_style('楷体_GB2312', 200, False, Halign=0, Valign=0, setBorder=0, setbgcolor=1))
        # 第7行
        sheet1.write_merge(6, 6, 0, 0, "支付给职工以及为职工支付的现金",
                           set_style('楷体_GB2312', 200, False, Halign=1, Valign=0, setBorder=0, setbgcolor=1))
        sheet1.write_merge(6, 6, 1, 1, 1, set_style('宋体', 200, False, Halign=0, Valign=0, setBorder=0, setbgcolor=1))
        # 第8行
        sheet1.write_merge(7, 7, 0, 0, "收到的税费返还",
                           set_style('楷体_GB2312', 200, False, Halign=1, Valign=0, setBorder=0, setbgcolor=1))
        sheet1.write_merge(7, 7, 1, 1, 2, set_style('宋体', 200, False, Halign=0, Valign=0, setBorder=0, setbgcolor=1))
        # 第9行
        sheet1.write_merge(8, 8, 0, 0, "收到其他与经营活动有关的现金",
                           set_style('楷体_GB2312', 200, False, Halign=1, Valign=0, setBorder=0, setbgcolor=1))
        sheet1.write_merge(8, 8, 1, 1, 3, set_style('宋体', 200, False, setbgcolor=1))
        # 第10行
        sheet1.write_merge(9, 9, 0, 0, "经营活动现金流入小计",
                           set_style('楷体_GB2312', 200, False, Halign=1, Valign=0, setBorder=0, setbgcolor=1))
        sheet1.write_merge(9, 9, 1, 1, 4, set_style('宋体', 200, False, Halign=0, Valign=0, setBorder=0, setbgcolor=1))
        # 第11行
        sheet1.write_merge(10, 10, 0, 0, "购买商品、接受劳务支付的现金",
                           set_style('楷体_GB2312', 200, False, Halign=1, Valign=0, setBorder=0, setbgcolor=1))
        sheet1.write_merge(10, 10, 1, 1, 5, set_style('宋体', 200, False, Halign=0, Valign=0, setBorder=0, setbgcolor=1))
        # 第12行
        sheet1.write_merge(11, 11, 0, 0, "支付给职工以及为职工支付的现金",
                           set_style('楷体_GB2312', 200, False, Halign=1, Valign=0, setBorder=0, setbgcolor=1))
        sheet1.write_merge(11, 11, 1, 1, 6, set_style('宋体', 200, False, Halign=0, Valign=0, setBorder=0, setbgcolor=1))
        # 第13行
        sheet1.write_merge(12, 12, 0, 0, "支付的各项税费",
                           set_style('楷体_GB2312', 200, False, Halign=1, Valign=0, setBorder=0, setbgcolor=1))
        sheet1.write_merge(12, 12, 1, 1, 7, set_style('宋体', 200, False, Halign=0, Valign=0, setBorder=0, setbgcolor=1))
        # 第14行
        sheet1.write_merge(13, 13, 0, 0, "支付其他与经营活动有关的现金",
                           set_style('楷体_GB2312', 200, False, Halign=1, Valign=0, setBorder=0, setbgcolor=1))
        sheet1.write_merge(13, 13, 1, 1, 8, set_style('宋体', 200, False, Halign=0, Valign=0, setBorder=0, setbgcolor=1))
        # 第15行
        sheet1.write_merge(14, 14, 0, 0, "经营活动现金流出小计",
                           set_style('楷体_GB2312', 200, False, Halign=1, Valign=0, setBorder=0, setbgcolor=1))
        sheet1.write_merge(14, 14, 1, 1, 9, set_style('宋体', 200, False, Halign=0, Valign=0, setBorder=0, setbgcolor=1))
        # 第16行
        sheet1.write_merge(15, 15, 0, 0, "经营活动产生的现金流量净额",
                           set_style('楷体_GB2312', 200, True, Halign=1, Valign=0, setBorder=0))
        sheet1.write_merge(15, 15, 1, 1, 10, set_style('宋体', 200, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
        # 第17行
        sheet1.write_merge(16, 16, 0, 0, "二、投资活动产生的现金流量：",
                           set_style('楷体_GB2312', 200, False, Halign=1, Valign=0, setBorder=0, setbgcolor=1))
        sheet1.write_merge(16, 16, 1, 1, "",
                           set_style('楷体_GB2312', 200, False, Halign=0, Valign=0, setBorder=0, setbgcolor=1))
        # 第18行
        sheet1.write_merge(17, 17, 0, 0, "收回投资收到的现金",
                           set_style('楷体_GB2312', 200, False, Halign=1, Valign=0, setBorder=0, setbgcolor=1))
        sheet1.write_merge(17, 17, 1, 1, 11, set_style('宋体', 200, False, Halign=0, Valign=0, setBorder=0, setbgcolor=1))
        # 第19行
        sheet1.write_merge(18, 18, 0, 0, "取得投资收益收到的现金",
                           set_style('楷体_GB2312', 200, False, Halign=1, Valign=0, setBorder=0, setbgcolor=1))
        sheet1.write_merge(18, 18, 1, 1, 12, set_style('宋体', 200, False, Halign=0, Valign=0, setBorder=0, setbgcolor=1))
        # 第20行
        sheet1.write_merge(19, 19, 0, 0, "处置固定资产、无形资产和其他长期资产收回的现金净额",
                           set_style('楷体_GB2312', 200, False, Halign=1, Valign=0, setBorder=0, setbgcolor=1))
        sheet1.write_merge(19, 19, 1, 1, 13, set_style('宋体', 200, False, Halign=0, Valign=0, setBorder=0, setbgcolor=1))
        # 第21行
        sheet1.write_merge(20, 20, 0, 0, "处置子公司及其他营业单位收到的现金净额",
                           set_style('楷体_GB2312', 200, False, Halign=1, Valign=0, setBorder=0, setbgcolor=1))
        sheet1.write_merge(20, 20, 1, 1, 14, set_style('宋体', 200, False, Halign=0, Valign=0, setBorder=0, setbgcolor=1))
        # 第22行
        sheet1.write_merge(21, 21, 0, 0, "收到其他与投资活动有关的现金",
                           set_style('楷体_GB2312', 200, False, Halign=1, Valign=0, setBorder=0, setbgcolor=1))
        sheet1.write_merge(21, 21, 1, 1, 15, set_style('宋体', 200, False, Halign=0, Valign=0, setBorder=0, setbgcolor=1))
        # 第23行
        sheet1.write_merge(22, 22, 0, 0, "投资活动现金流入小计",
                           set_style('楷体_GB2312', 200, False, Halign=1, Valign=0, setBorder=0, setbgcolor=1))
        sheet1.write_merge(22, 22, 1, 1, 16, set_style('宋体', 200, False, Halign=0, Valign=0, setBorder=0, setbgcolor=1))
        # 第24行
        sheet1.write_merge(23, 23, 0, 0, "购建固定资产、无形资产和其他长期资产支付的现金",
                           set_style('楷体_GB2312', 200, False, Halign=1, Valign=0, setBorder=0, setbgcolor=1))
        sheet1.write_merge(23, 23, 1, 1, 17, set_style('宋体', 200, False, Halign=0, Valign=0, setBorder=0, setbgcolor=1))
        # 第25行
        sheet1.write_merge(24, 24, 0, 0, "投资支付的现金",
                           set_style('楷体_GB2312', 200, False, Halign=1, Valign=0, setBorder=0, setbgcolor=1))
        sheet1.write_merge(24, 24, 1, 1, 18, set_style('宋体', 200, False, Halign=0, Valign=0, setBorder=0, setbgcolor=1))
        # 第26行
        sheet1.write_merge(25, 25, 0, 0, "取得子公司及其他营业单位支付的现金净额",
                           set_style('楷体_GB2312', 200, False, Halign=1, Valign=0, setBorder=0, setbgcolor=1))
        sheet1.write_merge(25, 25, 1, 1, 19, set_style('宋体', 200, False, Halign=0, Valign=0, setBorder=0, setbgcolor=1))
        # 第27行
        sheet1.write_merge(26, 26, 0, 0, "支付其他与投资活动有关的现金",
                           set_style('楷体_GB2312', 200, False, Halign=1, Valign=0, setBorder=0, setbgcolor=1))
        sheet1.write_merge(26, 26, 1, 1, 20, set_style('宋体', 200, False, Halign=0, Valign=0, setBorder=0, setbgcolor=1))
        # 第28行
        sheet1.write_merge(27, 27, 0, 0, "投资活动现金流出小计",
                           set_style('楷体_GB2312', 200, False, Halign=1, Valign=0, setBorder=0, setbgcolor=1))
        sheet1.write_merge(27, 27, 1, 1, 21, set_style('宋体', 200, False, Halign=0, Valign=0, setBorder=0, setbgcolor=1))
        # 第29行
        sheet1.write_merge(28, 28, 0, 0, "投资活动产生的现金流量净额",
                           set_style('楷体_GB2312', 200, True, Halign=1, Valign=0, setBorder=0, setbgcolor=0))
        sheet1.write_merge(28, 28, 1, 1, 22, set_style('宋体', 200, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
        # 第30行
        sheet1.write_merge(29, 29, 0, 0, "三、筹资活动产生的现金流量：",
                           set_style('楷体_GB2312', 200, False, Halign=1, Valign=0, setBorder=0, setbgcolor=1))
        sheet1.write_merge(29, 29, 1, 1, "",
                           set_style('楷体_GB2312', 200, False, Halign=0, Valign=0, setBorder=0, setbgcolor=1))
        # 第31行
        sheet1.write_merge(30, 30, 0, 0, "吸收投资收到的现金",
                           set_style('楷体_GB2312', 200, False, Halign=1, Valign=0, setBorder=0, setbgcolor=1))
        sheet1.write_merge(30, 30, 1, 1, 23, set_style('宋体', 200, False, Halign=0, Valign=0, setBorder=0, setbgcolor=1))
        # 第32行
        sheet1.write_merge(31, 31, 0, 0, "取得借款收到的现金",
                           set_style('楷体_GB2312', 200, False, Halign=1, Valign=0, setBorder=0, setbgcolor=1))
        sheet1.write_merge(31, 31, 1, 1, 24, set_style('宋体', 200, False, Halign=0, Valign=0, setBorder=0, setbgcolor=1))
        # 第33行
        sheet1.write_merge(32, 32, 0, 0, "收到其他与筹资活动有关的现金",
                           set_style('楷体_GB2312', 200, False, Halign=1, Valign=0, setBorder=0, setbgcolor=1))
        sheet1.write_merge(32, 32, 1, 1, 25, set_style('宋体', 200, False, Halign=0, Valign=0, setBorder=0, setbgcolor=1))
        # 第34行
        sheet1.write_merge(33, 33, 0, 0, "筹资活动现金流入小计",
                           set_style('楷体_GB2312', 200, False, Halign=1, Valign=0, setBorder=0, setbgcolor=1))
        sheet1.write_merge(33, 33, 1, 1, 26, set_style('宋体', 200, False, Halign=0, Valign=0, setBorder=0, setbgcolor=1))
        # 第35行
        sheet1.write_merge(34, 34, 0, 0, "偿还债务支付的现金",
                           set_style('楷体_GB2312', 200, False, Halign=1, Valign=0, setBorder=0, setbgcolor=1))
        sheet1.write_merge(34, 34, 1, 1, 27, set_style('宋体', 200, False, Halign=0, Valign=0, setBorder=0, setbgcolor=1))
        # 第36行
        sheet1.write_merge(35, 35, 0, 0, "分配股利、利润或偿付利息支付的现金",
                           set_style('楷体_GB2312', 200, False, Halign=1, Valign=0, setBorder=0, setbgcolor=1))
        sheet1.write_merge(35, 35, 1, 1, 28, set_style('宋体', 200, False, Halign=0, Valign=0, setBorder=0, setbgcolor=1))
        # 第37行
        sheet1.write_merge(36, 36, 0, 0, "支付其他与筹资活动有关的现金",
                           set_style('楷体_GB2312', 200, False, Halign=1, Valign=0, setBorder=0, setbgcolor=1))
        sheet1.write_merge(36, 36, 1, 1, 29, set_style('宋体', 200, False, Halign=0, Valign=0, setBorder=0, setbgcolor=1))
        # 第38行
        sheet1.write_merge(37, 37, 0, 0, "筹资活动现金流出小计",
                           set_style('楷体_GB2312', 200, False, Halign=1, Valign=0, setBorder=0, setbgcolor=1))
        sheet1.write_merge(37, 37, 1, 1, 30, set_style('宋体', 200, False, Halign=0, Valign=0, setBorder=0, setbgcolor=1))
        # 第39行
        sheet1.write_merge(38, 38, 0, 0, "筹资活动产生的现金流量净额",
                           set_style('楷体_GB2312', 200, True, Halign=1, Valign=0, setBorder=0, setbgcolor=0))
        sheet1.write_merge(38, 38, 1, 1, 31, set_style('宋体', 200, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
        # 第40行
        sheet1.write_merge(39, 39, 0, 0, "四、汇率变动对现金及现金等价物的影响",
                           set_style('楷体_GB2312', 200, False, Halign=1, Valign=0, setBorder=0, setbgcolor=1))
        sheet1.write_merge(39, 39, 1, 1, 32, set_style('宋体', 200, False, Halign=0, Valign=0, setBorder=0, setbgcolor=1))
        # 第41行
        sheet1.write_merge(40, 40, 0, 0, "五、现金及现金等价物净增加额",
                           set_style('楷体_GB2312', 200, True, Halign=1, Valign=0, setBorder=0, setbgcolor=0))
        sheet1.write_merge(40, 40, 1, 1, 33, set_style('宋体', 200, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
        # 第42行
        sheet1.write_merge(41, 41, 0, 0, "加：期初现金及现金等价物余额",
                           set_style('楷体_GB2312', 200, False, Halign=1, Valign=0, setBorder=0, setbgcolor=1))
        sheet1.write_merge(41, 41, 1, 1, "",
                           set_style('楷体_GB2312', 200, False, Halign=0, Valign=0, setBorder=0, setbgcolor=1))
        # 第43行
        sheet1.write_merge(42, 42, 0, 0, "六、期末现金及现金等价物余额",
                           set_style('楷体_GB2312', 200, False, Halign=1, Valign=0, setBorder=0, setbgcolor=1))
        sheet1.write_merge(42, 42, 1, 1, "",
                           set_style('楷体_GB2312', 200, False, Halign=0, Valign=0, setBorder=0, setbgcolor=1))
        # 第44行
        sheet1.write_merge(43, 43, 0, 0, "平衡",
                           set_style('楷体_GB2312', 200, False, Halign=1, Valign=0, setBorder=1, setbgcolor=1))

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

        filePath = 'C:/Users/Long/Desktop/ZFI77/saveFiles/'  # 保存文件路径
        fileName = 'ZFI77玉林现金流量表' + year + month.zfill(2)  # 保存文件名
        deleteOldFiles(filePath)  # 清空文件
        f.save(filePath + fileName + '.xls')  # 保存文件
        convFormat('C:\\Users\\Long\\Desktop\\ZFI77\\saveFiles\\',
                   'C:\\Users\\Long\\Desktop\\ZFI77\\')  # 转换表格格式(.xls -> .xlsx)
    except Exception as e:
        print("构造表或写入数据出错！！！", e)


if __name__ == '__main__':
    simulateMouseOperation()
    write_excel()
