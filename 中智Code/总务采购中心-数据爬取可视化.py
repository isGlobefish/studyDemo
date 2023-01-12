# -*- coding: utf-8 -*-
'''
@createTool    : PyCharm-2020.3.2
@projectName   : pythonProjectPy3.9
@originalAuthor: Made in win10.Sys design by deHao.Zou
@createTime    : 2021/02/12 10:11
'''
import os
import re
import cv2
import glob
import json
import time
import xlrd
import xlwt
import csv
import openpyxl
import datetime
import xlsxwriter
import pytesseract  # 用于图片转文字
import numpy as np
import pandas as pd
import urllib.request
from PIL import Image  # 用于打开图片和对图片处理
from termcolor import cprint
from selenium import webdriver
import win32com.client as win32
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC


def input(xpath, value):
    input_box = driver.find_element_by_xpath(xpath)
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


def switch(xpath):
    xf = driver.find_element_by_xpath(xpath)
    try:
        driver.switch_to.frame(xf)  # 切换
    except Exception as e:
        print('切换失败', e)


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
            print(">  已清空文件夹！！！")
        else:
            print("存在未删除文件")
    else:
        print("不存在excel文件")


# 合并数据
def mergeXlsTable(path):
    allXlsNum = glob.glob(path + "*.xls")
    print(">> 下载了" + str(len(allXlsNum)) + "个xls文件")

    def newDF():
        df = pd.DataFrame(
            columns=["日期", "产品名称", "市场", "生产企业", "规格型号", "数据类型", "最低价", "最高价", "平均价", "涨跌", "单位", "价格条件", "备注"])
        return df

    # 老百姓合并数据
    dataCGZX = newDF()
    newFileList = os.listdir(path)  # 该文件夹下所有的文件（包括文件夹）
    for refile in newFileList:
        refileName = os.path.splitext(refile)[0]  # 获取文件名
        refileType = os.path.splitext(refile)[1]  # 获取文件扩展名
        openFile = pd.read_excel(str(path) + str(refileName) + str(refileType), header=0)
        dataCGZX = dataCGZX.append(openFile)

    try:
        dataCGZX['日期'] = pd.to_datetime(dataCGZX['日期'], format='%Y/%m/%d').dt.date
        dataCGZX.to_excel('C:/Users/Zeus/Desktop/采购成本项目文件/卓创资讯MergeData.xlsx', index=False)
        print(">  数据合并保存输出成功！！！")
    except Exception as e:
        print(">  Fail合并数据！！！")


# def CGZX():  # 采购中心
global driver
print('\n' + ">>>数据爬取中,稍等片刻")
wei_zhi = "C:\\Users\\Zeus\\Desktop\\采购成本项目文件\\采购中心\\downloadFiles\\"  # 数据下载路径
deleteOldFiles(wei_zhi)  # 清空文件夹历史文件
options = webdriver.ChromeOptions()  # 打开设置
prefs = {'profile.default_content_settings.popups': 0, 'download.default_directory': wei_zhi}  # 设置路径
options.add_experimental_option('prefs', prefs)
driver = webdriver.Chrome(executable_path='C:\\Program Files (x86)\\Google\\Chrome\\Application\\chromedriver.exe', options=options)  # 打开浏览器
driver.implicitly_wait(30)  # 隐式等待
driver.get('https://www.sci99.com/')  # 进入网址
switch('/html/body/div[7]/div[1]/div[1]/iframe')  # 切入
input('/html/body/form/div[3]/ul/li[1]/input', 'zeuscgzx')  # 账号
input('/html/body/form/div[3]/ul/li[2]/input', 'zeuscgzx3737')  # 密码
button_a('/html/body/form/div[3]/div/input')  # 登录
driver.switch_to.parent_frame()  # 切出
time.sleep(10)

# 一、PVC粉
driver.get(
    'https://prices.sci99.com/cn/product_price.aspx?diid=6393&datatypeid=37&ppid=12349&ppname=PVC%u7C89&cycletype=day')  # 塑料—-通用塑料---PVC---PVC市场价格---宁波台塑
time.sleep(5)
driver.get(
    'https://prices.sci99.com/cn/product_price.aspx?diid=6393&datatypeid=37&ppid=12349&ppname=PVC%u7C89&cycletype=day')  # 塑料—-通用塑料---PVC---PVC市场价格---宁波台塑
time.sleep(5)
button_b('/html/body/form[1]/div[5]/div[5]/div/ul/li/a')  # 【PVC粉】导出数据
time.sleep(5)

# 二、玉米
driver.get(
    'https://prices.sci99.com/cn/product_price.aspx?diid=19561&datatypeid=331&ppid=12540&ppname=%u7389%u7C73&cycletype=day')  # 农产品---粮食网---玉米---玉米价格查询---蛇口港
time.sleep(5)
button_b('/html/body/form[1]/div[5]/div[5]/div/ul/li/a')  # 【玉米】导出数据
time.sleep(5)

# 三、白卡纸
driver.get(
    'https://prices.sci99.com/cn/product_price.aspx?diid=27734&datatypeid=37&ppid=12557&ppname=%u767D%u5361%u7EB8&cycletype=day')  # 林业---造纸网---包装用纸---包装用纸价格查询---白卡纸--华南
time.sleep(5)
button_b('/html/body/form[1]/div[5]/div[5]/div/ul/li/a')  # 【广西金桂---白卡纸】导出数据
time.sleep(5)
driver.get(
    'https://prices.sci99.com/cn/product_price.aspx?diid=12788&datatypeid=37&ppid=12557&ppname=%u767D%u5361%u7EB8&cycletype=day')  # 林业---造纸网---包装用纸---包装用纸价格查询---白卡纸--华南
time.sleep(5)
button_b('/html/body/form[1]/div[5]/div[5]/div/ul/li/a')  # 【宁波中华---白卡纸】导出数据
time.sleep(5)
driver.get(
    'https://prices.sci99.com/cn/product_price.aspx?diid=12860&datatypeid=37&ppid=12557&ppname=%u767D%u5361%u7EB8&cycletype=day')  # 林业---造纸网---包装用纸---包装用纸价格查询---白卡纸--华南
time.sleep(5)
button_b('/html/body/form[1]/div[5]/div[5]/div/ul/li/a')  # 【万国太阳---白卡纸】导出数据
time.sleep(5)
driver.get(
    'https://prices.sci99.com/cn/product_price.aspx?diid=12874&datatypeid=37&ppid=12557&ppname=%u767D%u5361%u7EB8&cycletype=day')  # 林业---造纸网---包装用纸---包装用纸价格查询---白卡纸--华南
time.sleep(5)
button_b('/html/body/form[1]/div[5]/div[5]/div/ul/li/a')  # 【江苏博汇---白卡纸】导出数据
time.sleep(5)

# 四、白板纸
driver.get(
    'https://prices.sci99.com/cn/product_price.aspx?diid=34463&datatypeid=37&ppid=12556&ppname=%u767D%u677F%u7EB8&cycletype=day')  # 林业---造纸网---包装用纸---包装用纸价格查询---白板纸--华南
time.sleep(5)
button_b('/html/body/form[1]/div[5]/div[5]/div/ul/li/a')  # 【东莞玖龙---白纸板】导出数据
time.sleep(5)
driver.get(
    'https://prices.sci99.com/cn/product_price.aspx?diid=12966&datatypeid=37&ppid=12556&ppname=%u767D%u677F%u7EB8&cycletype=day')  # 林业---造纸网---包装用纸---包装用纸价格查询---白板纸--华南
time.sleep(5)
button_b('/html/body/form[1]/div[5]/div[5]/div/ul/li/a')  # 【建晖纸业---白纸板】导出数据
time.sleep(5)

# 五、瓦楞纸
driver.get(
    'https://prices.sci99.com/cn/product_price.aspx?diid=112902&datatypeid=37&ppid=12573&ppname=%u74E6%u695E%u7EB8&cycletype=day')  # 林业---造纸网---包装用纸---华南---瓦楞纸---A级高瓦120g
time.sleep(5)
button_b('/html/body/form[1]/div[5]/div[5]/div/ul/li/a')  # 【瓦楞纸】导出数据
time.sleep(5)
driver.get(
    'https://prices.sci99.com/cn/product_price.aspx?diid=112903&datatypeid=37&ppid=12573&ppname=%u74E6%u695E%u7EB8&cycletype=day')  # 林业---造纸网---包装用纸---华南---瓦楞纸---AA级高瓦120g
time.sleep(5)
button_b('/html/body/form[1]/div[5]/div[5]/div/ul/li/a')  # 【瓦楞纸】导出数据
time.sleep(5)

# 六、箱板纸
driver.get(
    'https://prices.sci99.com/cn/product_price.aspx?diid=68005&datatypeid=37&ppid=12574&cycletype=day')  # 林业---造纸网---包装用纸---包装用纸价格查询---箱板纸---华东
time.sleep(5)
button_b('/html/body/form[1]/div[5]/div[5]/div/ul/li/a')  # 【AA级牛卡纸170g---箱板纸】导出数据
time.sleep(5)
driver.get(
    'https://prices.sci99.com/cn/product_price.aspx?diid=109779&datatypeid=37&ppid=12574&cycletype=day')  # 林业---造纸网---包装用纸---包装用纸价格查询---箱板纸---华南
time.sleep(5)
button_b('/html/body/form[1]/div[5]/div[5]/div/ul/li/a')  # 【A牛卡纸130克---箱板纸】导出数据
time.sleep(5)

# 七、镀锡板
driver.get(
    'https://prices.sci99.com/cn/product_price.aspx?diid=32764&datatypeid=37&ppid=12139&ppname=%u9540%u9521%u677F&cycletype=day')  # 钢铁---板材---更多---涂镀---镀锡板---镀锡专区（马口铁）---镀锡价格查询---宝钢集团
time.sleep(5)
button_b('/html/body/form[1]/div[5]/div[5]/div/ul/li/a')  # 【MR型 0.2*800-1000mm*C---镀锡板】导出数据
time.sleep(5)
driver.get(
    'https://prices.sci99.com/cn/product_price.aspx?diid=34525&datatypeid=37&ppid=12139&ppname=%u9540%u9521%u677F&cycletype=day')  # 钢铁---板材---更多---涂镀---镀锡板---镀锡专区（马口铁）---镀锡价格查询---宝钢集团
time.sleep(5)
button_b('/html/body/form[1]/div[5]/div[5]/div/ul/li/a')  # 【MR型 0.24*800-1000mm*C---镀锡板】导出数据
time.sleep(5)

# 八、白糖
driver.get(
    'https://prices.sci99.com/cn/product_price.aspx?diid=17968&datatypeid=37&ppid=12193&ppname=%u767D%u7CD6&cycletype=day')  # 农产品---糖业网---白糖---华南
time.sleep(5)
button_b('/html/body/form[1]/div[5]/div[5]/div/ul/li/a')  # 【白糖】导出数据
time.sleep(5)

mergeXlsTable('C:/Users/Zeus/Desktop/采购成本项目文件/采购中心/downloadFiles/')  # 合并数据
driver.implicitly_wait(30)
driver.quit()
