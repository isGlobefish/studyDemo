# -*- coding: utf-8 -*-
'''
@createTool    : PyCharm-2020.2.2
@projectName   : pythonProject 
@originalAuthor: Made in win10.Sys design by deHao.Zou
@createTime    : 2020/11/20 17:58
'''
import os
import re
import cv2
import glob
import json
import time
import xlrd
import xlwt
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

today = int(time.strftime("%d", time.localtime()))
if today <= 2:
    lastMonth = "2"  # 上个月数据下载
else:
    toMonth = "1"  # 本月数据下载


def newdfLBX():
    dflbx = pd.DataFrame(
        columns=["区域", "月份", "日期", "商品编码", "商品名称", "单位", "规格", "数量", "单价", "金额", "业务部门", "厂家", "年份", "区域目标"])
    return dflbx


def input_b(id_1, value):
    input_box = driver.find_element_by_xpath(id_1)
    try:
        input_box.send_keys(value)
    except Exception as e:
        print('fail', e)


def button_a(xpath):
    global wait
    wait = WebDriverWait(driver, 3)
    button = driver.find_element_by_xpath(xpath)
    try:
        button.click()
    except Exception as e:
        print('fail搜索', e)


def button_b(xpath):
    global wait
    wait = WebDriverWait(driver, 3)
    button = driver.find_element_by_xpath(xpath)
    try:
        driver.execute_script("$(arguments[0]).click()", button)
    except Exception as e:
        print('fail搜索', e)


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


def time_0(time):
    if time > 9:
        time1 = str(time)
    else:
        time1 = '0' + str(time)
    return time1


def last_day_of_month(any_day):
    next_month = any_day.replace(day=28) + datetime.timedelta(days=4)
    return next_month - datetime.timedelta(days=next_month.day)


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
def convFormat(openPath, savePath, finalPath, dkhName):
    fileList = os.listdir(openPath)  # 该文件夹下所有的文件（包括文件夹）
    print("转换" + str(fileList) + "文件格式")
    for file in fileList:  # 遍历所有文件
        fileName = os.path.splitext(file)[0]  # 获取文件名
        fileType = os.path.splitext(file)[1]  # 获取文件扩展名
        openFiles = openPath + fileName + fileType
        saveFiles = savePath + fileName + fileType
        finalFiles = finalPath + dkhName + str(month) + fileType
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        wb = excel.Workbooks.Open(openFiles)
        wb.SaveAs(saveFiles + "x", FileFormat=51)  # FileFormat = 51转为.xlsx、FileFormat = 56转为.xls
        wb.SaveAs(finalFiles + "x", FileFormat=51)  # FileFormat = 51转为.xlsx、FileFormat = 56转为.xls
        wb.Close()
        excel.Application.Quit()
    print(dkhName + "（xls->xlsx）已转换格式完成！！！")


# 合并老百姓
def mergeXlsTable():
    global toMonth
    path = "E:/大客户数据/LBX/mergeTableLBX/"
    allXlsNum = glob.glob(path + "*.xls")
    print("下载了" + str(len(allXlsNum)) + "个xls文件")

    def mergeTableLBX():
        # 新建一个dataframe文件
        dfLBX = pd.DataFrame(
            columns=["区域", "月份", "日期", "商品编码", "商品名称", "单位", "规格", "数量", "单价", "金额", "业务部门", "厂家", "年份", "区域目标"])
        return dfLBX

    # 排序文件并命名
    oldFileList = os.listdir(path)  # 该文件夹下所有的文件（包括文件夹）
    createTimeList = []
    for oldfile in oldFileList:
        fileName = os.path.splitext(oldfile)[0]  # 获取文件名
        createTime = int(fileName[20:34])  # 获取时间部分数字
        createTimeList.append(createTime)  # 合并为一个列表
    createTimeListSort = sorted(createTimeList, reverse=False)  # reverse=False升序【从小到大，也即是先下载的在最前面，也即是1号到N号】
    # 按下载顺序（日期顺序）重命名文件
    for indexfile, valuefile in enumerate(createTimeListSort):  # 下角标：start = 1
        for file in oldFileList:  # 遍历所有文件
            '''
            if os.path.isdir(Olddir):   #如果是文件夹则跳过
                continue
            '''
            Olddir = os.path.join(path, file)  # 原来的文件路径
            fileMergeName = 'storeProductSummary_' + str(valuefile)
            indexFile = indexfile + 1
            fileName = os.path.splitext(file)[0]  # 获取文件名
            fileType = os.path.splitext(file)[1]  # 获取文件扩展名
            fileNewName = str(year) + '-' + str(month).zfill(2) + '-' + str(indexFile).zfill(2)
            if fileMergeName == fileName:
                Newdir = os.path.join(path, fileNewName + fileType)
                os.rename(Olddir, Newdir)  # 文件重命名
    # 老百姓合并数据
    dataLBX = mergeTableLBX()
    newFileList = os.listdir(path)  # 该文件夹下所有的文件（包括文件夹）
    for refile in newFileList:
        refileName = os.path.splitext(refile)[0]  # 获取文件名
        refileType = os.path.splitext(refile)[1]  # 获取文件扩展名
        openFile = pd.read_excel(str(path) + str(refileName) + str(refileType))
        openFile.insert(0, "日期", '')
        openFile["日期"] = refileName
        dataLBX = dataLBX.append(openFile)
    if toMonth == '1':
        dataLBX.to_excel(
            'E:/大客户数据/LBX/dataLBX' + str(year) + '-' + str(month).zfill(2) + '-' + str(day).zfill(2) + '.xlsx',
            index=False)
    else:
        dataLBX.to_excel(
            'E:/大客户数据/LBX/dataLBX' + str(year) + '-' + str(month).zfill(2) + '-' + str(lastDay).zfill(2) + '.xlsx',
            index=False)
    print("老百姓数据合并输出完成！！！")


if toMonth == "1":  # 判断时间
    year = int(time.strftime("%Y", time.localtime()))  # 本年
    month = int(time.strftime("%m", time.localtime()))  # 本月
    day = int(time.strftime("%d", time.localtime()))  # 本日
    time0 = str(year) + "-" + time_0(month) + "-" + str('01')  # 本月1号
    time15 = str(year) + "-" + time_0(month) + "-" + str('15')  # 本月15号
    time16 = str(year) + "-" + time_0(month) + "-" + str('16')  # 本月16号
    time1 = str(year) + "-" + time_0(month) + "-" + time_0(day - 1)  # 本月前一天
    time2 = str(year) + "-" + time_0(month) + "-" + time_0(day - 2)  # 本月前两天
    lasttime = last_day_of_month(datetime.date(year, month, day))
else:
    year = int(time.strftime("%Y", time.localtime()))  # 本年
    month = int(time.strftime("%m", time.localtime())) - 1  # 上一月
    day = int(time.strftime("%d", time.localtime()))  # 本日
    time0 = str(year) + "-" + time_0(month) + "-" + str('01')  # 本月1号
    time15 = str(year) + "-" + time_0(month) + "-" + str('15')  # 本月15号
    time16 = str(year) + "-" + time_0(month) + "-" + str('16')  # 本月16号
    lasttime = last_day_of_month(datetime.date(year, month, day))
    lastDay = int(str(lasttime)[8:10])  # 上个月最后一天
    time1 = str(lasttime)
    time2 = str(lasttime)


def LBX():  # 老百姓
    global driver
    print('\n' + "开始导出老百姓数据ing......")
    wei_zhi = "E:\\大客户数据\\LBX\\mergeTableLBX\\"  # 老百姓下载路径
    deleteOldFiles(wei_zhi)  # 清空老百姓文件夹文件
    options = webdriver.ChromeOptions()  # 打开设置
    prefs = {'profile.default_content_settings.popups': 0, 'download.default_directory': wei_zhi}  # 设置路径
    options.add_experimental_option('prefs', prefs)
    driver = webdriver.Chrome(executable_path='C:\\Program Files (x86)\\Google\\Chrome\\Application\\chromedriver.exe',
                              options=options)  # 打开浏览器
    driver.implicitly_wait(30)  # 隐式等待
    driver.get('http://srm.lbxdrugs.com:6066/srm/coframe/auth/login/login.jsp')  # 进入网址
    input_b('/html/body/div/div[1]/div/form/div[1]/span/span/input', '900022800005')  # 账号
    input_b('/html/body/div/div[1]/div/form/div[2]/span/span/input', 'zz106126')  # 密码
    button_a('/html/body/div/div[1]/div/form/div[4]/input')  # 登录
    time.sleep(3)
    # button_b('/html/body/div[2]/div/div[5]/span')
    button_a('/html/body/div[2]/div/div[2]/div/div/div[1]/div[1]/ul/li[1]/dl/dt')  # 销售管理
    time.sleep(0.5)
    button_a('/html/body/div[2]/div/div[2]/div/div/div[1]/div[1]/ul/li[1]/dl/dd[1]/ul/li/a')  # 门店销售明细
    time.sleep(3)
    switch('/html/body/div[2]/div/div[2]/div/div/div[2]/div[2]/div/div/iframe')  # 切换明细
    input_b('/html/body/div[1]/fieldset/table/tbody/tr[1]/td[2]/span/span/input', '')  # 激活文本输入框
    time.sleep(0.5)
    xuanfu('/html/body/div[1]/fieldset/table/tbody/tr[1]/td[2]/span')
    button_b('/html/body/div[1]/fieldset/table/tbody/tr[1]/td[2]/span')  # 鼠标点击省公司
    time.sleep(2)
    button_b('/html/body/div[7]/div/div[1]/div[1]/table/tbody/tr/td[1]/input')  # 全选
    button_b('/html/body/div[1]/fieldset/table/tbody/tr[1]/td[2]/span/span/span/span[2]/span')  # 鼠标点击省公司
    time.sleep(0.5)
    input_b('/html/body/div[1]/fieldset/table/tbody/tr[1]/td[4]/span/span/input', '')  # 激活文本输入框
    time.sleep(0.5)
    button_b('/html/body/div[1]/fieldset/table/tbody/tr[1]/td[4]/span/span/span/span[2]/span')  # 商品名称
    driver.switch_to.parent_frame()  # 切出
    switch('/html/body/div[4]/div/div[2]/div[2]/iframe')  # 切换
    time.sleep(5)
    button_b('/html/body/div[3]/div/div[2]/div[2]/div[2]/table/tbody/tr[2]/td[3]/div/div[1]/input')  # 双击
    time.sleep(1)
    button_b('/html/body/div[3]/div/div[2]/div[2]/div[2]/table/tbody/tr[2]/td[3]/div/div[1]/input')  # 全选
    time.sleep(1)
    button_b('/html/body/a[1]/span')  # 确定
    driver.switch_to.parent_frame()  # 切出
    if toMonth == "1":
        for dayi in range(1, day):
            lbsTime = str(year) + "-" + str(month).zfill(2) + "-" + str(dayi).zfill(2)
            switch('/html/body/div[2]/div/div[2]/div/div/div[2]/div[2]/div/div/iframe')  # 切换明细
            driver.find_element_by_xpath(
                '/html/body/div[1]/fieldset/table/tbody/tr[2]/td[2]/span[1]/span[1]/input').clear()  # 清空开始时间
            time.sleep(1)
            input_b('/html/body/div[1]/fieldset/table/tbody/tr[2]/td[2]/span[1]/span[1]/input', lbsTime)  # 输入开始日期
            time.sleep(0.5)
            driver.find_element_by_xpath(
                '/html/body/div[1]/fieldset/table/tbody/tr[2]/td[2]/span[3]/span[1]/input').clear()  # 清空结尾时间
            time.sleep(1)
            input_b('/html/body/div[1]/fieldset/table/tbody/tr[2]/td[2]/span[3]/span[1]/input', lbsTime)  # 输入日期
            time.sleep(0.5)
            button_b('/html/body/div[2]/a[1]/span')  # 查询
            time.sleep(2)
            button_b('/html/body/div[2]/a[2]/span')  # 导出
            driver.switch_to.parent_frame()  # 切出
            time.sleep(5)
    else:
        for dayi in range(1, lastDay + 1):
            lbsTime = str(year) + "-" + str(month).zfill(2) + "-" + str(dayi).zfill(2)
            switch('/html/body/div[2]/div/div[2]/div/div/div[2]/div[2]/div/div/iframe')  # 切换明细
            driver.find_element_by_xpath(
                '/html/body/div[1]/fieldset/table/tbody/tr[2]/td[2]/span[1]/span[1]/input').clear()  # 清空开始时间
            time.sleep(1)
            input_b('/html/body/div[1]/fieldset/table/tbody/tr[2]/td[2]/span[1]/span[1]/input', lbsTime)  # 输入开始日期
            time.sleep(0.5)
            driver.find_element_by_xpath(
                '/html/body/div[1]/fieldset/table/tbody/tr[2]/td[2]/span[3]/span[1]/input').clear()  # 清空结尾时间
            time.sleep(1)
            input_b('/html/body/div[1]/fieldset/table/tbody/tr[2]/td[2]/span[3]/span[1]/input', lbsTime)  # 输入日期
            time.sleep(0.5)
            button_b('/html/body/div[2]/a[1]/span')  # 查询
            time.sleep(2)
            button_b('/html/body/div[2]/a[2]/span')  # 导出
            driver.switch_to.parent_frame()  # 切出
            time.sleep(5)
    mergeXlsTable()  # 合并表格


try:
    LBX()  # 老百姓
except Exception as e:
    print("老百姓导出出错", e)

if toMonth == '1':
    openData = pd.read_excel(
        'E:/大客户数据/LBX/dataLBX' + str(year) + '-' + str(month).zfill(2) + '-' + str(day).zfill(2) + '.xlsx',
        sheet_name=0, header=0, index_col=None)
else:
    openData = pd.read_excel(
        'E:/大客户数据/LBX/dataLBX' + str(year) + '-' + str(month).zfill(2) + '-' + str(lastDay).zfill(2) + '.xlsx',
        sheet_name=0, header=0, index_col=None)

lbxdf = newdfLBX()  # 老百姓
lbxdf["业务部门"] = openData["业务部门"]
lbxdf["区域"] = lbxdf.apply(lambda x: dictMatchStr1.setdefault(str(x["业务部门"]), 'NA'), axis=1)
lbxdf["月份"] = str(month) + "月"
lbxdf["日期"] = openData["日期"]
lbxdf["商品编码"] = openData["商品编码"]
lbxdf["商品名称"] = openData["商品名称"]
lbxdf["单位"] = openData["单位"]
lbxdf["规格"] = openData["规格"]
lbxdf["数量"] = openData["数量"]
lbxdf["单价"] = lbxdf.apply(lambda x: dictMatchFloat.setdefault(float(x["商品编码"]), 0), axis=1)
lbxdf["金额"] = lbxdf["数量"].map(float) * lbxdf["单价"].map(float)
lbxdf["厂家"] = openData["厂家"]
lbxdf["年份"] = year
lbxdf["区域目标"] = lbxdf.apply(lambda x: dictMatchStr2.setdefault(str(x["业务部门"]), 'NA'), axis=1)

lbxdf.to_excel('C:/Users/Long/Desktop/老百姓2020年11月1-24日.xlsx', index=0)

import xlrd

# 区域、区域目标匹配字典
dictMatchStr1 = {}
dictMatchStr2 = {}
dictStr = xlrd.open_workbook('C:/Users/Long/Desktop/老百姓sheet2自动填充表.xlsx')  # 载入字典
tableStr = dictStr.sheet_by_name('Sheet1')
rowStr = tableStr.nrows
for i in range(1, rowStr):
    colValStr = tableStr.row_values(i)
    dictMatchStr1[str(colValStr[0])] = colValStr[1]
    dictMatchStr2[str(colValStr[0])] = colValStr[2]

# 单价匹配字典
dictMatchFloat = {}
dictFloat = xlrd.open_workbook('Z:/龙展华/DKH/商品编码对照字典.xlsx')  # 载入字典
tableFloat = dictFloat.sheet_by_name('Sheet1')
rowFloat = tableFloat.nrows
for i in range(1, rowFloat):
    colValFloat = tableFloat.row_values(i)
    dictMatchFloat[float(colValFloat[1])] = colValFloat[8]

openDataMatch = openData.drop_duplicates(subset=["业务部门"], keep='first')
openDataRetIndex = openDataMatch.reset_index(drop=True)
readDataMatch = pd.read_excel('C:/Users/Long/Desktop/老百姓sheet2自动填充表.xlsx', header=0)
shopNameList = readDataMatch["业务部门"].drop_duplicates().values.tolist()



outData = newdfLBX()
for i in range(len(openDataRetIndex)):
    if openDataRetIndex.at[i, "业务部门"] not in shopNameList:
        outData = outData.append(openDataRetIndex.loc[[i]])
outData.to_excel("C:/Users/Long/Desktop/新增店名.xlsx")


dd = readDataMatch.append(outData["业务部门"])
result = pd.concat([readDataMatch, outData], axis=1, sort=False)










import pandas as pd
import xlrd
workbook = xlrd.open_workbook('C:/Users/Long/Desktop/04-玉林现金流量表2019年1月.xlsx')  # 载入字典
work = workbook.sheet_by_index(0)
print(work.merged_cells)

work.to_excel('C:/Users/Long/Desktop/123.xlsx')
work.save('C:/Users/Long/Desktop/123.xlsx')




data = pd.read_excel("C:/Users/Long/Desktop/04-玉林现金流量表2019年1月.xlsx")
data.to_excel("C:/Users/Long/Desktop/123.xlsx")

# row0 = [u'业务', u'状态', u'北京', u'上海', u'广州', u'深圳', u'状态小计', u'合计']
# column0 = [u'机票', u'船票', u'火车票', u'汽车票', u'其它']
# status = [u'预订', u'出票', u'退票', u'业务小计']
#
# # 生成第一行
# for i in range(0, len(row0)):
#     sheet1.write(0, i, row0[i], set_style('Times New Roman', 220, True))
#
# # 生成第一列和最后一列(合并4行)
# i, j = 1, 0
# while i < 4 * len(column0) and j < len(column0):
#     sheet1.write_merge(i, i + 3, 0, 0, column0[j], set_style('Arial', 220, True))  # 第一列
#     sheet1.write_merge(i, i + 3, 7, 7)  # 最后一列"合计"
#     i += 4
#     j += 1
#
# sheet1.write_merge(21, 21, 0, 1, u'合计', set_style('Times New Roman', 220, True))
#
# # 生成第二列
# i = 0
# while i < 4 * len(column0):
#     for j in range(0, len(status)):
#         sheet1.write(j + i + 1, 1, status[j])
#     i += 4














