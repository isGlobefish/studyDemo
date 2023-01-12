# -*- coding: utf-8 -*-
'''
@createTool    : PyCharm-2020.2.2
@projectName   : pythonProjectPy3.9
@originalAuthor: Made in win10.Sys design by deHao.Zou
@createTime    : 2020/12/11 8:28
'''
import os
import re
import csv
import cv2
import json
import time
import xlrd
import xlwt
import glob
import openpyxl
import datetime
import xlsxwriter
import numpy as np
import pandas as pd
import urllib.request
from termcolor import cprint, colored
from selenium import webdriver
import win32com.client as win32
import pytesseract  # 用于图片转文字
from PIL import Image  # 用于打开图片和对图片处理
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC

aaa = "1"  # 下载本月数据

def get_exce(wei_zhi):
    all_exce = glob.glob(wei_zhi + "*.xls")
    print("该目录下有" + str(len(all_exce)) + "个excel文件：")
    if (len(all_exce) == 0):
        return 0
    else:
        for i in range(len(all_exce)):
            print(all_exce[i])
        return all_exce

def input_a(id_1, value):
    input_box = driver.find_element_by_id(id_1)
    try:
        input_box.send_keys(value)
    except Exception as e:
        print('fail输入', e)


def input_b(id_1, value):
    input_box = driver.find_element_by_xpath(id_1)
    try:
        input_box.send_keys(value)
    except Exception as e:
        print('fail输入', e)


def button_a(xpath):
    global wait
    wait = WebDriverWait(driver, 3)
    button = driver.find_element_by_xpath(xpath)
    try:
        button.click()
    except Exception as e:
        print('fail点击', e)


def button_b(xpath):
    global wait
    wait = WebDriverWait(driver, 3)
    button = driver.find_element_by_xpath(xpath)
    try:
        driver.execute_script("$(arguments[0]).click()", button)
    except Exception as e:
        print('fail点击', e)


def clear_a(id_1):
    input_box = driver.find_element_by_xpath(id_1)
    try:
        input_box.clear()
    except Exception as e:
        print('fail清空输入框', e)


def switch(xpath):
    xf = driver.find_element_by_xpath(xpath)
    try:
        driver.switch_to.frame(xf)  # 切换
    except Exception as e:
        print('切换失败', e)


def xuanfu(xpath):
    element = driver.find_element_by_xpath(xpath)
    try:
        ActionChains(driver).move_to_element(element).perform()  # 鼠标悬浮
    except Exception as e:
        print('悬浮失败', e)


# 获取excel文件下的所有sheet
def get_sheet(fh):
    sheets = fh.sheets()
    return sheets


# 获取sheet下有多少行数据
def get_sheetrow_num(sheet):
    return sheet.nrows


def get_sheet_data(sheet, row, j):
    for i in range(row - 1):
        if (i == 0):
            global biao_tou
            biao_tou = ['类别', '商品SAP编码', '商品名称', '规格', '单位', '店号/区域ID', '店名/区域', '销量', '过账日期', '合同价', '省份']
            continue
        values = sheet.row_values(i)
        values.append(sf[j])
        all_data1.append(values)
    return all_data1


def time_0(time):
    if time > 9:
        time1 = str(time)
    else:
        time1 = '0' + str(time)
    return time1


def last_day_of_month(any_day):
    next_month = any_day.replace(day=28) + datetime.timedelta(days=4)
    return next_month - datetime.timedelta(days=next_month.day)


if aaa == "1":  # 判断时间
    year = int(time.strftime("%Y", time.localtime()))  # 年
    month = int(time.strftime("%m", time.localtime()))  # 月
    day = int(time.strftime("%d", time.localtime()))  # 日
    daySub1 = day - 1
    time0 = str(year) + "-" + time_0(month) + "-" + str('01')
    time15 = str(year) + "-" + time_0(month) + "-" + str('15')
    time16 = str(year) + "-" + time_0(month) + "-" + str('16')
    time1 = str(year) + "-" + time_0(month) + "-" + time_0(day - 1)
    time2 = str(year) + "-" + time_0(month) + "-" + time_0(day - 2)
    lasttime = last_day_of_month(datetime.date(year, month, day))
    lastDay = int(day)
    dateDSL = time0 + ' - ' + time2  # 本月大参林下载日期范围
else:
    year = int(time.strftime("%Y", time.localtime()))  # 年
    # year = 2021
    month = int(time.strftime("%m", time.localtime())) - 1  # 上月
    # month = 12
    day = int(time.strftime("%d", time.localtime()))  # 日
    # day = 31
    time0 = str(year) + "-" + time_0(month) + "-" + str('01')
    time15 = str(year) + "-" + time_0(month) + "-" + str('15')
    time16 = str(year) + "-" + time_0(month) + "-" + str('16')
    lasttime = last_day_of_month(datetime.date(year, month, day))
    lastDay = int(str(lasttime)[8:10])  # 上个月最后一天
    dateDSL = time0 + ' - ' + str(lasttime)  # 上月大参林下载日期范围
    time1 = str(lasttime)
    time2 = str(lasttime)


def HW():  # 海王
    print('\n' + ">>>【海王】数据爬取中,稍等片刻")
    global sf
    global driver
    global all_data1
    all_data1 = []
    sf = ['北京', '长春', '成都', '大连', '电商', '福州', '广州', '河南', '湖北', '湖南', '深圳总部',
          '杭州', '江苏', '辽宁', '宁波', '青岛', '潍坊', '上海', '沈阳', '深圳', '天津', '泰州']  # 海王省份列表
    wei_zhi = "D:\\FilesCenter\\大客户数据\\HW\\"  # 海王下载路径
    # szLocation = 'D:/FilesCenter/大客户数据/HW-深圳/'
    options = webdriver.ChromeOptions()  # 打开设置
    prefs = {
        'profile.default_content_settings.popups': 0,
        'download.default_directory': wei_zhi
    }  # 设置路径
    options.add_experimental_option('prefs', prefs)
    driver = webdriver.Chrome(executable_path='C:\\Program Files (x86)\\Google\\Chrome\\Application\\chromedriver.exe', options=options)  # 打开浏览器
    driver.implicitly_wait(5)  # 隐式等待
    with open('D://FilesCenter//大客户数据//anquan.txt', 'r', encoding='utf8') as f:
        listCookies = json.loads(f.read())  # 读取cookies
    driver.get('http://srm.nepstar.cn')  # 进入登录界面
    driver.delete_all_cookies()  # 删旧cookies
    for cookie in listCookies:
        if 'expiry' in cookie:
            del cookie['expiry']
        driver.add_cookie(cookie)  # 新增cookies
    time.sleep(3)
    driver.get('http://srm.nepstar.cn/ELSServer_HWXC/default2.jsp?account=110829_1001&loginChage=N&telphone1=18279409642')
    # 读取完cookie刷新页面
    time.sleep(3)
    button_a('//*[@id="treeMenu"]/li[7]/a')  # 数据查询
    time.sleep(0.5)
    button_a('//*[@id="salesOutSourcingInfoManage"]')  # 销售查询
    time.sleep(0.5)
    # xf = driver.find_element_by_xpath('/html/body/div[2]/div/nav[2]/div[3]/iframe')  # 先通过xpath定位到iframe
    xf = driver.find_element_by_xpath('/html/body/div[2]/div/nav[2]/div[2]/iframe')  # 先通过xpath定位到iframe
    driver.switch_to.frame(xf)  # 再将定位对象传给switch_to.frame()方法
    button_a('/html/body/div[1]/div[2]/form/div[1]/div/div/span')  # 商品编码
    time.sleep(5)
    driver.switch_to.parent_frame()
    # xf = driver.find_element_by_xpath('/html/body/div[4]/div/table/tbody/tr[2]/td/div/iframe')
    xf = driver.find_element_by_xpath('/html/body/div[5]/div/table/tbody/tr[2]/td/div/iframe')
    driver.switch_to.frame(xf)  # 切换
    button_a('/html/body/div/main/div[1]/div[1]/div[1]/table/thead/tr/th[2]/div/span/input')  # 全选
    time.sleep(0.5)
    button_a('/html/body/div/main/div[2]/button[1]')  # 确定
    time.sleep(0.5)
    driver.switch_to.parent_frame()
    xf = driver.find_element_by_xpath('/html/body/div[2]/div/nav[2]/div[2]/iframe')
    driver.switch_to.frame(xf)  # 切换
    input_b('/html/body/div[1]/div[2]/form/div[2]/div/div/input', time0)  # 查询日期
    time.sleep(0.5)
    input_b('/html/body/div[1]/div[2]/form/div[3]/div/div/input', time2)  # 延迟两天
    time.sleep(0.5)
    button_a('/html/body/div[1]/div[2]/form/div[4]/div/div/div/p')  # 联采合同
    time.sleep(0.5)
    button_a('/html/body/div[1]/div[2]/form/div[4]/div/div/div/div/ul/li[3]')
    time.sleep(0.5)
    for i in range(2, 24):
        if i == 12:
            cprint("跳过深圳总部", 'cyan', attrs=['bold', 'reverse', 'blink'])
        # elif i == 15:
        #     cprint("无权限访问，跳过辽宁", 'cyan', attrs=['bold', 'reverse', 'blink'])
        else:
            button_a('/html/body/div[1]/div[2]/form/div[5]/div/div/div/p')  # 选地区
            time.sleep(0.5)
            ix = '/html/body/div[1]/div[2]/form/div[5]/div/div/div/div/ul/li[' + str(i) + ']'
            button_a(ix)  # 选好地区
            time.sleep(0.5)
            button_a('/html/body/div[1]/div[2]/form/button[2]')  # 查询
            time.sleep(7)
            button_a('/html/body/div[1]/div[3]/div[1]/div[1]/table/thead/tr/th[2]/div/span')  # 全选
            time.sleep(1)
            button_a('/html/body/div[1]/div[1]/nav/div/div/ul/li[1]/a')  # 查看明细
            time.sleep(1)
            driver.switch_to.parent_frame()  # 回退
            try:
                # driver.find_element_by_xpath('/html/body/div[1]/div[1]/nav/div/div/ul/li[1]/a')  # 查看明细是否存在
                xf = driver.find_element_by_xpath('/html/body/div[2]/div/nav[2]/div[3]/iframe')  # 切入明细
                driver.switch_to.frame(xf)
                time.sleep(5)
                button_a('/html/body/div/div[1]/nav/div/div/ul/li[2]/a')  # 导出
                time.sleep(6)
                element = driver.find_element_by_xpath('/html/body/div/div[1]/nav/div/div/ul/li[2]')
                ActionChains(driver).move_to_element(element).perform()  # 鼠标悬浮
                button_a('/html/body/div/div[1]/nav/div/div/ul/li[1]/a')  # 返回
                driver.switch_to.parent_frame()  # 切出
                time.sleep(8)
                all_exce = get_exce(wei_zhi)
                # 得到要合并的所有exce表格数据
                if (all_exce == 0):
                    cprint(sf[i - 2] + "下载出错！！！", 'magenta', attrs=['bold', 'reverse', 'blink'])
                else:
                    for exce in all_exce:
                        fh = xlrd.open_workbook(exce)
                        # 打开文件
                        sheets = get_sheet(fh)
                        # 获取文件下的sheet数量
                        for sheet in range(len(sheets)):
                            row = get_sheetrow_num(sheets[sheet])
                            # 获取一个sheet下的所有的数据的行数
                            all_data1 = get_sheet_data(sheets[sheet], row,i - 2)
                            os.remove(exce)
                print("导出完成！", sf[i - 2])
                # xf = driver.find_element_by_xpath('/html/body/div[2]/div/nav[2]/div[3]/iframe')  # 切入查询
                xf = driver.find_element_by_xpath('/html/body/div[2]/div/nav[2]/div[2]/iframe')  # 切入查询
                time.sleep(1)
                driver.switch_to.frame(xf)  # 切换
                button_a('/html/body/div[1]/div[3]/div[1]/div[1]/table/thead/tr/th[2]/div/span')  # 全不选

            except BaseException:
                cprint(sf[i - 2] + "->列无数据", 'magenta', attrs=['bold', 'reverse', 'blink'])
                xf = driver.find_element_by_xpath('/html/body/div[2]/div/nav[2]/div[2]/iframe')  # 切入查询
                driver.switch_to.frame(xf)  # 切换
                button_a('/html/body/div[1]/div[3]/div[1]/div[1]/table/thead/tr/th[2]/div/span')  # 全不选
            # finally:

    dictCookies = driver.get_cookies()  # 获取cookies
    jsonCookies = json.dumps(dictCookies)
    with open('D://FilesCenter//大客户数据//anquan.txt', 'w') as f:
        f.write(jsonCookies)  # 保存新cookies

    biao_tou = ['类别', '商品SAP编码', '商品名称', '规格', '单位', '店号/区域ID', '店名/区域', '销量', '过账日期', '合同价', '省份']
    all_data1.insert(0, biao_tou)  # 表头写入
    # 下面开始文件数据的写入
    new_excel = "D:\\FilesCenter\\大客户数据\\HW\\" + "HAIWANG" + str(month) + ".xlsx"  # 新建的excel文件名字
    fh1 = xlsxwriter.Workbook(new_excel)  # 新建一个excel表
    new_sheet = fh1.add_worksheet()  # 新建一个sheet表
    for i in range(len(all_data1)):
        for j in range(len(all_data1[i])):
            c = all_data1[i][j]
            new_sheet.write(i, j, c)
    fh1.close()  # 关闭该excel表


try:
    HW()  # 海王
except Exception as e:
    print("海王导出出错", e)
