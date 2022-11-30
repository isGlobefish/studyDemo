# -*- coding: utf-8 -*-
"""
Created on Sat May  9 18:06:26 2020

@author: Long
"""

import os
import json
import time
import xlrd
import requests
from PIL import Image
from selenium import webdriver
from qiniu import Auth, put_file, etag
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains


# 登陆报表
def checkin_get_brower(i0116, i0118):
    brower = webdriver.Chrome(
        executable_path=r'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chromedriver.exe')  # 打开谷歌
    brower.maximize_window()  # 全屏
    brower.get(url)  # 进入网站
    time.sleep(2)
    brower.find_element_by_id('i0116').send_keys(i0116)  # 账号
    brower.find_element_by_id('idSIButton9').click()  # 下一步
    time.sleep(2)
    brower.find_element_by_id('i0118').send_keys(i0118)  # 密码
    brower.find_element_by_id('idSIButton9').click()  # 登陆
    time.sleep(2)
    a = brower.find_element_by_xpath(
        "/html/body/div/form/div[1]/div/div[1]/div[2]/div/div[2]/div/div[3]/div[2]/div/div/div[1]")  # 否
    a.click()
    time.sleep(10)
    return brower


# 截图
def jietu(xpath):
    global local_file
    while True:
        try:
            displayArea = brower.find_element_by_xpath(xpath)  # 检查截图页面是否存在
        except:
            print('等待刷新10秒')
            time.sleep(5)
        else:
            break
    time.sleep(3)
    left = displayArea.location['x']  # x起点
    top = displayArea.location['y']  # y起点
    right = displayArea.location['x'] + displayArea.size['width']  # 宽
    bottom = displayArea.location['y'] + displayArea.size['height']  # 高
    brower.get_screenshot_as_file('screenshot.png')  # 截图
    im = Image.open('screenshot.png')  # 选择图片
    im = im.crop((left, top, right, bottom))  # 对浏览器截图进行裁剪
    im.save(local_file)  # 保存


# 上传图片
def get_img_url():
    global key
    access_key = "DZnCErimkn2yQrn4aYel3JX7vPXKRonlvDFoVh1e"
    secret_key = "FBEHIFyMG28nWZrn316df-ny5bmIz_LanRWtabCi"
    q = Auth(access_key, secret_key)
    bucket_name = "qiniu730173201"  # 网上的空间
    token = q.upload_token(bucket_name, key)  # 删掉旧图片
    ret, info = put_file(token, key, local_file)  # 上传新图片
    base_url = "http://zzsy.zeus.cn/"  # 二级域名
    url = base_url + '/' + key
    private_url = q.private_download_url(url)  # 图片新网址
    # r = requests.get(private_url)
    # assert r.status_code == 200` </pre>
    return private_url


def dingmessage(str1, img_url0, ddurl):  # 上传dingding
    webhook = ddurl  # 机器人url
    header = {
        "Content-Type": "application/json",
        "Charset": "UTF-8"
    }
    message = {
        "msgtype": "markdown",
        "markdown": {"title": "报表",
                     "text": "# " + str1 + "\n\n" +
                             "#### 每日报表推送\n\n" +
                             "> ![screenshot](" + img_url0 + ")\n" +
                             "> ###### 由数据运营中心发布\n"
                     },
        "at": {"isAtAll": 0}
    }  # 报表内容
    message_json = json.dumps(message)  # json格式化
    info = requests.post(url=webhook, data=message_json, headers=header)  # 机器人发送信息
    print(info.text)


def button_a(xpath):
    global wait
    wait = WebDriverWait(brower, 3)
    button = brower.find_element_by_xpath(xpath)  # 点击方式
    try:
        button.click()
    except Exception as e:
        print('fail搜索', e)


def button_b(xpath):
    global wait
    wait = WebDriverWait(brower, 3)
    button = brower.find_element_by_xpath(xpath)
    try:
        brower.execute_script("$(arguments[0]).click()", button)  # 点击方式
    except Exception as e:
        print('fail搜索', e)


def switch(xpath):
    xf = brower.find_element_by_xpath(xpath)
    try:
        brower.switch_to.frame(xf)  # 切换页面
    except Exception as e:
        print('切换失败', e)


def xuanfu(xpath):
    element = brower.find_element_by_xpath(xpath)
    try:
        ActionChains(brower).move_to_element(element).perform()  # 鼠标悬浮
    except Exception as e:
        print('悬浮失败', e)


def dd(x1, x2):
    for i in range(x1, x2):
        values = table.row_values(i)
        print(values[0])
        button_a('//span[contains(@title,"主要指标")]')
        time.sleep(1)

        # button_a('//div[contains(@aria-label,"办事处")]')
        # win32api.mouse_event(win32con.MOUSEEVENTF_WHEEL,0,0,-1)
        # xuanfu('/html/body/div[10]/div[1]/div/div[2]/div/div[3]/div/div[3]')
        button_a('//div[contains(@aria-label,"' + values[0] + '")]')
        time.sleep(0.5)
        # button_a('//div[contains(@aria-label,"办事处")]')

        filename = 'changzhou'
        local_file = 'E:\\钉钉群机器人发送\\displayArea' + filename + '.png'  # 本地存放路径
        if os.path.exists(local_file):
            os.remove(local_file)  # 删除旧的
        else:
            print('no such file')
        jietu('//div[contains(@class,"displayArea disableAnimations fitToPage")]')
        # jietu('/html/body/div[1]/root-downgrade/mat-sidenav-container/mat-sidenav-content/div/landing/div/div/div/ng-transclude/landing-route/report/exploration-container/exploration-container-modern/div/div/exploration-host/div/div/exploration/div/explore-canvas-modern/div/div[2]/div/div[2]/div[2]')
        # 截图
        global key
        img_name = 'displayArea' + filename + '.png'  # 网络命名
        key = 'data/%s' % (img_name)  # 网络路径
        img_url = get_img_url()  # 上传下载
        time.sleep(1)
        # 钉钉发送
        dingmessage(values[0], img_url, values[1])

        button_a('//span[contains(@title,"概况")]')
        if os.path.exists(local_file):
            os.remove(local_file)  # 删除旧的
        else:
            print('no such file')
        jietu('//div[contains(@class,"displayArea disableAnimations fitToPage")]')
        img_url = get_img_url()  # 上传下载
        time.sleep(1)
        # 钉钉发送
        dingmessage(values[0], img_url, values[1])

        #        button_a('//span[contains(@title,"客户概况")]')
        #        if os.path.exists(local_file):
        #            os.remove(local_file)#删除旧的
        #        else:
        #            print('no such file')
        #        jietu('//div[contains(@class,"displayArea disableAnimations fitToPage")]')
        #        img_url = get_img_url()#上传下载
        #        time.sleep(1)
        #        #钉钉发送
        #        dingmessage(values[0],img_url,values[1])

        button_a('//span[contains(@title,"办事处状况")]')
        if os.path.exists(local_file):
            os.remove(local_file)  # 删除旧的
        else:
            print('no such file')
        jietu('//div[contains(@class,"displayArea disableAnimations fitToPage")]')
        img_url = get_img_url()  # 上传下载
        time.sleep(1)
        # 钉钉发送
        dingmessage(values[0], img_url, values[1])

        button_a('//span[contains(@title,"1页")]')
        if os.path.exists(local_file):
            os.remove(local_file)  # 删除旧的
        else:
            print('no such file')
        jietu('//div[contains(@class,"displayArea disableAnimations fitToPage")]')
        img_url = get_img_url()  # 上传下载
        time.sleep(1)
        # 钉钉发送
        dingmessage(values[0], img_url, values[1])

        button_a('//span[contains(@title,"2页")]')
        if os.path.exists(local_file):
            os.remove(local_file)  # 删除旧的
        else:
            print('no such file')
        jietu('//div[contains(@class,"displayArea disableAnimations fitToPage")]')
        img_url = get_img_url()  # 上传下载
        time.sleep(1)
        # 钉钉发送
        dingmessage(values[0], img_url, values[1])


file_path = r'E:/数据/DD机器人/机器人编号1.xlsx'
data = xlrd.open_workbook(file_path)
sheets = data.sheets()  # 读取机器人编号
row = sheets[0].nrows
table = sheets[0]

url = "https://app.powerbi.cn/groups/me/reports/f0d56bbc-5bd2-4799-8f82-188ccf68b425?ctid=5f382c82-06e2-4960-bb1e-def38b69931b"
brower = checkin_get_brower('chenjia@zzeus.partner.onmschina.cn', 'Zeus0760')  # 登陆
time.sleep(100)  # 等待数据下载完成
dd(1, 26)  # 开始运行26个
# dd(1,2)
brower.quit()
