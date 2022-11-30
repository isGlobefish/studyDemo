# -*- coding: utf-8 -*-
'''
@createTool    : PyCharm-2020.2.2
@projectName   : pythonProjectPy3.9 
@originalAuthor: Made in win10.Sys design by deHao.Zou
@createTime    : 2020/12/22 15:13

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
    print('\n结束,按任意键退出....')  # 检测到Ctrl+c组合键结束运行

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
ap.ShellExecute(0, 'open', 'D:\\Software_manage\\SAP\\SAPgui\\saplogon.exe', '', '', 1)  # 打开SAP Logon
time.sleep(6)
pg.press('down', presses=2, interval=0.125)  # 按下 2次
time.sleep(1)
pg.press('enter')  # 回车
time.sleep(3)
user = '106610'  # 用户账号
password = 'abc123'  # 密码
cp.copy(user)  # 复制用户账号
pg.hotkey('ctrl', 'v')  # 粘贴用户账号
time.sleep(1)
pg.press('down')  # 按下
cp.copy(password)  # 复制密码
pg.hotkey('ctrl', 'v')  # 粘贴密码
pg.press('enter')  # 回车
time.sleep(4)

# ========== MB5L ======================
affcode1 = 'MB5L'  # 事务编码
# pg.typewrite(affcode1)
cp.copy(affcode1)  # 复制文件路径
pg.hotkey('ctrl', 'v')  # 粘贴路径
pg.press('enter')  # 回车
time.sleep(3)
pg.press('down')  # 按下
time.sleep(1)
compcode1 = '2000'  # 公司代码
cp.copy(compcode1)  # 复制公司代码
pg.hotkey('ctrl', 'v')  # 粘贴公司代码
pg.press('down', presses=4, interval=0.1)  # 按down 4次
pg.press('tab', presses=3, interval=0.1)  # 按tab 3次
pg.press('down', presses=1, interval=0.1)  # 按down 1次 选择余额,前期
time.sleep(2)
pg.press('F8')  # 按F8
time.sleep(10)
pg.click(56, 22, clicks=1)  # 清单
time.sleep(2)
pg.press('s', presses=1, interval=0.5)
time.sleep(1)
pg.press('F', presses=1, interval=0.5)
time.sleep(1)
pg.press('down', presses=1, interval=0.1)  # 按down 1次
pg.press('enter')  # 回车
time.sleep(2)
pg.press('up')
time.sleep(1)
pg.hotkey('ctrl', 'A')  # 全选
pg.press('backspace', presses=1, interval=0.5)
# pg.press('end')
# pg.press('backspace', presses=50, interval=0.01)
time.sleep(1)
path = 'C:\\Users\\Zeus\\Desktop\\AA-核对库存科目金额\\MB5L\\'
cp.copy(path)  # 复制文件路径
pg.hotkey('ctrl', 'v')  # 粘贴路径
time.sleep(1)
pg.press('down', presses=1, interval=0.1)  # 按down 1次
# pg.press('tab', presses=1, interval=0.1)  # 按tab 1次
pg.hotkey('ctrl', 'A')  # 全选
pg.press('backspace', presses=1, interval=0.5)
# pg.press('backspace', presses=4, interval=0.01)
MB5LFileName = 'MB5L' + year + month.zfill(2) + '.xls'
cp.copy(MB5LFileName)  # 复制文件路径
pg.hotkey('ctrl', 'v')  # 粘贴路径
time.sleep(0.5)
pg.press('enter')  # 回车
time.sleep(1)
pg.press('left')  # 左
pg.press('enter')  # 回车
time.sleep(2)

# ========== ZFI64 ======================
pg.press('F3', presses=2, interval=0.5)  # 后退
affcode2 = 'ZFI64'  # 事务编码
# pg.typewrite(affcode2)
cp.copy(affcode2)
pg.hotkey('ctrl', 'v')  # 粘贴
pg.press('enter')  # 回车
time.sleep(2)
pg.hotkey('ctrl', 'A')  # 全选
pg.press('backspace', presses=1, interval=0.5)
cp.copy(year)
pg.hotkey('ctrl', 'v')
time.sleep(1)
pg.press('down', presses=1, interval=0.1)
cp.copy(month)
pg.hotkey('ctrl', 'v')
time.sleep(1)
pg.press('down', presses=1, interval=0.1)
factoryStart = '2000'
cp.copy(factoryStart)  # 复制开始工厂
pg.hotkey('ctrl', 'v')  # 粘贴
pg.press('tab', presses=1)
time.sleep(0.5)
factoryEnd = '2001'
cp.copy(factoryEnd)  # 复制最后工厂
pg.hotkey('ctrl', 'v')  # 粘贴
pg.press('F8')
time.sleep(350)
# 获取当前屏幕分辨率
screenWidth, screenHeight = pg.size()
pg.click(screenWidth / 2, screenHeight / 2, clicks=1, button='right')  # 右键
time.sleep(1)
pg.press('down', presses=8, interval=0.1)  # 按下 8次
pg.press('enter', presses=2, interval=1)  # 回车继续
time.sleep(10)
pg.press('tab', presses=6, interval=0.1)
time.sleep(1)
pg.press('down', presses=1, interval=1)
time.sleep(1)
pg.press('space', presses=1, interval=1)  # 选中桌面
time.sleep(1)
pg.press('tab', presses=5, interval=0.1)  # 选中第一个文件
time.sleep(1)
pg.press('down', presses=8, interval=0.1)  # 按下 8次
pg.press('enter', presses=1)  # 回车继续
time.sleep(1)
pg.press('down', presses=3, interval=0.1)  # 按下 3次
pg.press('enter', presses=1)  # 回车继续
time.sleep(1)
pg.press('tab', presses=2, interval=0.1)
pg.press('backspace', presses=1, interval=0.5)
ZFI64fileName = 'ZFI64' + year + month.zfill(2) + '.xlsx'
cp.copy(ZFI64fileName)  # 复制文件路径
pg.hotkey('ctrl', 'v')  # 粘贴路径
pg.press('enter')  # 保存文件
time.sleep(2)
pg.press('left', presses=1, interval=1)
pg.press('enter', presses=1, interval=1)
time.sleep(2)
pg.press('left', presses=1, interval=1)
pg.press('enter', presses=1, interval=1)
time.sleep(10)
pg.press(1578, 10, presses=1)  # 关闭表格
time.sleep(2)
pg.press('F3', presses=2, interval=0.5)
time.sleep(1)

# ========== ZMM45 ======================
affcode3 = 'ZMM45'  # 事务编码
cp.copy(affcode3)
pg.hotkey('ctrl', 'v')  # 粘贴
pg.press('enter')  # 回车
time.sleep(3)
compcode2 = '2000'
cp.copy(compcode2)
pg.hotkey('ctrl', 'v')  # 粘贴
pg.press('down', presses=2, interval=0.1)
pg.hotkey('ctrl', 'A')  # 全选
pg.press('backspace', presses=1, interval=0.5)
cp.copy(year)
pg.hotkey('ctrl', 'v')
time.sleep(1)
pg.press('down', presses=1, interval=0.1)
cp.copy(month)
pg.hotkey('ctrl', 'v')
time.sleep(1)
pg.press('F8')
time.sleep(150)
# 获取当前屏幕分辨率
screenWidth, screenHeight = pg.size()
pg.click(screenWidth / 2, screenHeight / 2, clicks=1, button='right')  # 右键
time.sleep(1)
pg.press('down', presses=6, interval=0.1)  # 按下 6次
pg.press('enter', presses=1, interval=1)  # 回车继续
time.sleep(25)
pg.press('tab', presses=6, interval=0.1)
time.sleep(1)
pg.press('down', presses=1, interval=1)
time.sleep(1)
pg.press('space', presses=1, interval=1)  # 选中桌面
time.sleep(1)
pg.press('tab', presses=1, interval=1)  # 选中第一个文件
pg.press('down', presses=8, interval=0.1)  # 按下 8次
pg.press('enter', presses=1)  # 回车继续
time.sleep(1)
pg.press('down', presses=4, interval=0.1)
pg.press('enter', presses=1)  # 选中文件夹
time.sleep(1)
pg.press('tab', presses=2, interval=1)  # 全选文件名
pg.press('backspace', presses=1, interval=0.5)  # 删除文件名
time.sleep(1)
ZFI64FileName = 'ZMM45' + year + month.zfill(2) + '.xlsx'
cp.copy(ZFI64FileName)
pg.hotkey('ctrl', 'v')  # 粘贴
pg.press('enter', presses=1)  # 保存文件
time.sleep(2)
pg.press('left', presses=1, interval=1)
pg.press('enter', presses=1, interval=1)
time.sleep(2)
pg.press('left', presses=1, interval=1)
pg.press('enter', presses=1, interval=1)
time.sleep(10)
pg.press(1578, 10, presses=1)  # 关闭表格
time.sleep(2)
pg.press('F3', presses=2, interval=0.5)
time.sleep(1)

# ========== ZFI32 ======================
affcode3 = 'ZFI32'  # 事务编码
cp.copy(affcode3)
pg.hotkey('ctrl', 'v')  # 粘贴
pg.press('enter')  # 回车
time.sleep(3)
pg.press('down', presses=1, interval=0.1)
btYear = '2020-011'
cp.copy(btYear)
pg.hotkey('ctrl', 'v')  # 粘贴
pg.press('tab', presses=5, interval=0.1)
time.sleep(2)
copy14 = """
1403010100
1403020100
1403030100
1404010100
1405010100
1405010200
1405010300
1405010400
1405020100
1405020200
1405020300
1405020400
1405029900
1411010100
"""
cp.copy(copy14)
pg.hotkey('shift', 'F12')  # 全选
time.sleep(2)
pg.press('F8', presses=2, interval=2)
time.sleep(20)
# 获取当前屏幕分辨率
screenWidth, screenHeight = pg.size()
pg.click(screenWidth * 3 / 8, screenHeight * 3 / 8, clicks=1, button='right')  # 右键
time.sleep(1)
pg.press('down', presses=5, interval=0.1)  # 按下 5次
pg.press('enter', presses=1, interval=1)  # 回车继续
time.sleep(5)
pg.press('tab', presses=6, interval=0.1)
time.sleep(1)
pg.press('down', presses=1, interval=1)
time.sleep(1)
pg.press('space', presses=1, interval=1)  # 选中桌面
time.sleep(1)
pg.press('tab', presses=1, interval=1)  # 选中第一个文件
pg.press('down', presses=8, interval=0.1)  # 按下 8次
pg.press('enter', presses=1)  # 回车继续
time.sleep(1)
pg.press('down', presses=1, interval=0.1)
pg.press('enter', presses=1)  # 选中文件夹
time.sleep(1)
pg.press('tab', presses=2, interval=1)  # 全选文件名
pg.press('backspace', presses=1, interval=0.5)  # 删除文件名
time.sleep(1)
ZFI32FileName = 'ZFI32' + year + month.zfill(2) + '.xlsx'
cp.copy(ZFI32FileName)
pg.hotkey('ctrl', 'v')  # 粘贴
pg.press('enter', presses=1)  # 保存文件
time.sleep(2)
pg.press('left', presses=1, interval=1)
pg.press('enter', presses=1, interval=1)
time.sleep(2)
pg.press('left', presses=1, interval=1)
pg.press('enter', presses=1, interval=1)
time.sleep(10)
pg.press(1578, 10, presses=1)  # 关闭表格
