# -*- coding: utf-8 -*-
'''
@createTool    : PyCharm-2020.2.2
@projectName   : PythonCode 
@originalAuthor: Made in 14471 design by deHao.Zou
@createTime    : 2020/10/10 11:06
'''
import glob
import os
import cv2
import re
import json
import time
import datetime
import xlrd
import xlsxwriter
import pandas as pd
import numpy as np
import urllib.request
from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from PIL import Image  # 用于打开图片和对图片处理
import pytesseract  # 用于图片转文字


def input_a(id_1,value):
    input_box = driver.find_element_by_id(id_1)
    try:
        input_box.send_keys(value)
    except Exception as e:
        print('fail',e)

def input_b(id_1,value):
    input_box = driver.find_element_by_xpath(id_1)
    try:
        input_box.send_keys(value)
    except Exception as e:
        print('fail',e)

def button_a(xpath):
    global wait
    wait = WebDriverWait(driver, 3)
    button = driver.find_element_by_xpath(xpath)
    try:
        button.click()
    except Exception as e:
        print('fail搜索',e)

def button_b(xpath):
    global wait
    wait = WebDriverWait(driver, 3)
    button = driver.find_element_by_xpath(xpath)
    try:
        driver.execute_script("$(arguments[0]).click()",button)
    except Exception as e:
        print('fail搜索',e)

def clear_a(id_1):
    input_box = driver.find_element_by_xpath(id_1)
    try:
        input_box.clear()
    except Exception as e:
        print('fail清空输入框',e)

def switch(xpath):
    xf = driver.find_element_by_xpath(xpath)
    try:
        driver.switch_to.frame(xf)#切换
    except Exception as e:
        print('切换失败',e)

def xuanfu(xpath):
    element = driver.find_element_by_xpath(xpath)
    try:
        ActionChains(driver).move_to_element(element).perform()#鼠标悬浮
    except Exception as e:
        print('悬浮失败',e)

def time_0(time):
    if time>9:
        time1 = str(time)
    else:
        time1 = '0' +str(time)
    return time1

def last_day_of_month(any_day):
    next_month = any_day.replace(day=28) + datetime.timedelta(days=4)
    return next_month - datetime.timedelta(days=next_month.day)

aaa = "1"

if aaa == "1": #判断时间
    year = int(time.strftime("%Y", time.localtime()))  # 年
    month = int(time.strftime("%m", time.localtime()))   # 月
    day = int(time.strftime("%d", time.localtime()))  # 日
    time0 = str(year) + "-" + time_0(month) + "-" + str('01')
    time15 = str(year) + "-" + time_0(month) + "-" + str('15')
    time16 = str(year) + "-" + time_0(month) + "-" + str('16')
    time1 = str(year) + "-" + time_0(month) + "-" + time_0(day-1)
    time2 = str(year) + "-" + time_0(month) + "-" + time_0(day-2)
    lasttime = last_day_of_month(datetime.date(year, month, day))
else:
    year = int(time.strftime("%Y", time.localtime()))  # 年
    month = int(time.strftime("%m", time.localtime()))-1   # 月
    day = int(time.strftime("%d", time.localtime()))  # 日
    time0 = str(year) + "-" + time_0(month) + "-" + str('01')
    time15 = str(year) + "-" + time_0(month) + "-" + str('15')
    time16 = str(year) + "-" + time_0(month) + "-" + str('16')
    lasttime = last_day_of_month(datetime.date(year, month, day))
    lastDay = int(str(lasttime)[8:10])  # 上个月最后一天
    time1 = str(lasttime)
    time2 = str(lasttime)

def GD():
    global driver  # 将driver变量定义为全局变量
    wei_zhi = "E:\\大客户数据\\GD\\"  # 文件保存路径
    options = webdriver.ChromeOptions()  # 打开设置
    prefs = {'profile.default_content_settings.popups': 0,
             'download.default_directory': wei_zhi}  # 设置谷歌保存形式：0代表不弹出窗口下载
    options.add_experimental_option('prefs', prefs)
    driver = webdriver.Chrome(
        executable_path='C:\\Program Files (x86)\\Google\\Chrome\\Application\\chromedriver.exe',
        options=options)  # 打开浏览器
    driver.get('http://221.133.237.227:801')  # 打开登陆页面
    input_b('//input[@id="loginId"]', '601317')  # 账号
    input_b('//input[@id="password"]', 'zb888888')  # 密码
    cishu = 5  # 改为5次10次太久
    while True:
        try:
            if cishu > 0:
                try:
                    # 获取验证码
                    driver.save_screenshot('picture.png')  # 全屏截图
                    page_snap_obj = Image.open('picture.png')
                    img = driver.find_element_by_xpath('/html/body/form/div/div[2]/p[4]/span[2]/img')  # 验证码元素位置
                    time.sleep(1)
                    location = img.location
                    size = img.size  # 获取验证码的大小参数
                    left = location['x']
                    top = location['y']
                    right = left + size['width']
                    bottom = top + size['height']
                    image_obj = page_snap_obj.crop((left, top, right, bottom))  # 按照验证码的长宽，切割验证码
                    img = image_obj.convert("L")  # 转灰度
                    pixdata = img.load()
                    w, h = img.size
                    threshold = 160
                    # 遍历所有像素，大于阈值的为黑色
                    for y in range(h):
                        for x in range(w):
                            if pixdata[x, y] < threshold:
                                pixdata[x, y] = 0
                            else:
                                pixdata[x, y] = 255
                    data = img.getdata()
                    w, h = img.size
                    black_point = 0
                    for x in range(1, w - 1):
                        for y in range(1, h - 1):
                            mid_pixel = data[w * y + x]  # 中央像素点像素值
                            if mid_pixel < 50:  # 找出上下左右四个方向像素点像素值
                                top_pixel = data[w * (y - 1) + x]
                                left_pixel = data[w * y + (x - 1)]
                                down_pixel = data[w * (y + 1) + x]
                                right_pixel = data[w * y + (x + 1)]
                                # 判断上下左右的黑色像素点总个数
                                if top_pixel < 10:
                                    black_point += 1
                                if left_pixel < 10:
                                    black_point += 1
                                if down_pixel < 10:
                                    black_point += 1
                                if right_pixel < 10:
                                    black_point += 1
                                if black_point < 1:
                                    img.putpixel((x, y), 255)
                                black_point = 0
                    img.save('image.png')
                    pytesseract.pytesseract.tesseract_cmd = r"E:\tesseract-v5.0.0\tesseract.exe"  # 设置pyteseract路径
                    result = pytesseract.image_to_string(Image.open('image.png'))  # 识别图片里面的文字
                    dropSpecialString = re.sub(u"([^\u4e00-\u9fa5\u0030-\u0039\u0041-\u005a\u0061-\u007a])", "",
                                               result)  # 去除识别出来的特殊字符
                    num = dropSpecialString[0:4]  # 只获取前4个字符
                    print(num)
                    input_b('//input[@id="imagecode"]', num)  # 验证码
                    time.sleep(0.5)
                    button_a('//*[@onclick="SumitLogin();"]')  # 登陆
                    time.sleep(5)
                    button_a('/html/body/div[2]/div[2]/div/ul/li/ul/li[5]/div/span[4]')  # 门店零售
                    break
                except:
                    print('登陆失败', cishu)
                    clear_a('//input[@id="imagecode"]')  # 清空文本框
                    cishu = cishu - 1
            else:
                break
        except:
            pass

    button_a('/html/body/div[2]/div[2]/div/ul/li/ul/li[5]/div/span[3]')  # 门店零售
    time.sleep(2)
    #driver.switch_to.parent_frame()#切出
    #button_a('/html/body/form/fieldset/table/tbody/tr[1]/td[1]/span/span/span')#开始日期
    #time.sleep(0.5)
    #button_a('//*[@abbr="'+ str(year) +','+ str(month) +',1"]')
    #button_a('/html/body/form/fieldset/table/tbody/tr[1]/td[1]/span/span/span')#日期
    #time.sleep(0.5)
    #button_a('//*[@abbr="'+ str(year) +','+ str(month) +','+ str(day-1) +'"]')
    time.sleep(2)
    switch('/html/body/div[3]/div/div/div[2]/div[2]/div/iframe')  # 切入
    if aaa == "1": #判断时间
        js_begin = 'document.querySelector("#searchArea > tbody > tr:nth-child(1) > td:nth-child(1) > span > input.combo-text.validatebox-text").removeAttribute("readonly");'
        driver.execute_script(js_begin)
        # 用js方法输入日期
        js_begin1_value = 'document.querySelector("#searchArea > tbody > tr:nth-child(1) > td:nth-child(1) > span > input.combo-text.validatebox-text").value="' +  time0 +'"'  #表面输入本月第一天
        driver.execute_script(js_begin1_value)
        js_begin2_value = 'document.querySelector("#searchArea > tbody > tr:nth-child(1) > td:nth-child(1) > span > input.combo-value").value="' +  time0 +'"'  #实际有效的输入本月第一天
        driver.execute_script(js_begin2_value)
        time.sleep(2)
        js_end = 'document.querySelector("#searchArea > tbody > tr:nth-child(1) > td:nth-child(2) > span > input.combo-text.validatebox-text").removeAttribute("readonly");'
        driver.execute_script(js_end)
        js_end1_value = 'document.querySelector("#searchArea > tbody > tr:nth-child(1) > td:nth-child(2) > span > input.combo-text.validatebox-text").value="' + time1 +'"'  # 表面输入前一天
        driver.execute_script(js_end1_value)
        js_end2_value = 'document.querySelector("#searchArea > tbody > tr:nth-child(1) > td:nth-child(2) > span > input.combo-value").value="' + time1 +'"'  # 实际有效的输入前一天
        driver.execute_script(js_end2_value)
    else:
        js_begin = 'document.querySelector("#searchArea > tbody > tr:nth-child(1) > td:nth-child(1) > span > input.combo-text.validatebox-text").removeAttribute("readonly");'
        driver.execute_script(js_begin)
        # 用js方法输入日期
        js_begin1_value = 'document.querySelector("#searchArea > tbody > tr:nth-child(1) > td:nth-child(1) > span > input.combo-text.validatebox-text").value="' +  time0 +'"'  #表面输入上月第一天
        driver.execute_script(js_begin1_value)
        js_begin2_value = 'document.querySelector("#searchArea > tbody > tr:nth-child(1) > td:nth-child(1) > span > input.combo-value").value="' +  time0 +'"'  #实际有效的输入上月第一天
        driver.execute_script(js_begin2_value)
        time.sleep(2)
        js_end = 'document.querySelector("#searchArea > tbody > tr:nth-child(1) > td:nth-child(2) > span > input.combo-text.validatebox-text").removeAttribute("readonly");'
        driver.execute_script(js_end)
        js_end1_value = 'document.querySelector("#searchArea > tbody > tr:nth-child(1) > td:nth-child(2) > span > input.combo-text.validatebox-text").value="' + time1 +'"'  # 表面输入上个月最后一天
        driver.execute_script(js_end1_value)
        js_end2_value = 'document.querySelector("#searchArea > tbody > tr:nth-child(1) > td:nth-child(2) > span > input.combo-value").value="' + time1 +'"'  #实际有效的输入上个月最后一天
        driver.execute_script(js_end2_value)
    time.sleep(2)
    # button_a('/html/body/div[1]/div/div[1]/table/tbody/tr/td[1]/a/span/span')
    # print("稍等片刻......")
    # time.sleep(30)
    button_a('/html/body/div[1]/div/div[1]/table/tbody/tr/td[3]/a/span/span')  # 导出数据
    time.sleep(3)
    print("输出成功！！！")
    time.sleep(1)
    driver.switch_to.parent_frame() # 切出
    time.sleep(30)


GD()
