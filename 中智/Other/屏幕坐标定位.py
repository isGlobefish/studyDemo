# -*- coding: utf-8 -*-
'''
@createTool    : PyCharm-2020.2.2
@projectName   : pythonCode 
@originalAuthor: Made in win10.Sys design by deHao.Zouxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
@createTime    : 2020/10/12 14:06
'''

'''

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


def simulateMouseOperation():
    try:
        ap.ShellExecute(0, 'open', 'D:\\SAP\\SAPgui\\saplogon.exe', '', '', 1)  # 打开SAP Logon
        time.sleep(5)
        pg.press('down', presses=2, interval=0.25)  # 按下 2次
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
        affcode = 'ZFI77'  # 事务编码
        cp.copy(affcode)  # 复制事务编码
        pg.hotkey('ctrl', 'v')  # 粘贴事务编码
        pg.press('enter')  # 回车
        time.sleep(3)
        compcode = '2200'  # 公司代码
        cp.copy(compcode)  # 复制公司代码
        pg.hotkey('ctrl', 'v')  # 粘贴公司代码
        pg.press('down')  # 按下
        time.sleep(1)
        finversion = '20180101'  # 财务版本
        cp.copy(finversion)  # 复制财务版本
        pg.hotkey('ctrl', 'v')  # 粘贴财务版本
        pg.press('down')  # 按下
        time.sleep(1)
        finyear = '2019'  # 财务年度
        pg.press('backspace', presses=4, interval=0.125)  # 按backspace 4次
        cp.copy(finyear)  # 复制财务年度
        pg.hotkey('ctrl', 'v')  # 粘贴财务年度
        pg.press('down')  # 按下
        pg.press('enter')  # 回车
        time.sleep(2)
        pg.press('down')  # 按下
        finmonth = '10'  # 财务期间
        pg.press('backspace',presses=2,interval=0.125)  # 按backspace 2次
        cp.copy(finmonth)  # 复制财务期间
        pg.hotkey('ctrl', 'v')  # 粘贴财务期间
        pg.press('F8')  # 按F8
        time.sleep(3)
        pg.click(59, 18, clicks=1)  # 点击列表
        time.sleep(1)
        pg.click(84, 90, clicks=1)  # 点击导出
        time.sleep(1)
        pg.click(302, 110, clicks=1)  # 点击电子表格
        time.sleep(2)
        pg.press('enter')  # 回车继续
        time.sleep(2)
        pg.click(291, 199, clicks=1)  # 点击桌面
        time.sleep(2)
        pg.click(436, 143, clicks=1)  # 点击第一个文件
        time.sleep(2)
        pg.press('down', presses=8, interval=0.25)  # 按下 8次
        time.sleep(2)
        pg.press('enter')  # 回车
        time.sleep(2)
        pg.press('down', presses=1, interval=1)  # 按下
        time.sleep(2)
        pg.press('enter', presses=1, interval=2)  # 回车 1次
        time.sleep(2)
        pg.press('tab', presses=4, interval=0.125)  # tab 4次
        pg.click(749, 442, clicks=1)  # 取消
        time.sleep(2)
        pg.click(1114, 8, clicks=1)  # 关闭
        time.sleep(2)
        pg.click(236, 364, clicks=1)  # 是
        time.sleep(2)
        pg.click(1450, 13, clicks=1)  # 退出
    except Exception as e:
        print("模拟鼠标操作出错，请修改出错的定位参数！！！", e)


simulateMouseOperation()
