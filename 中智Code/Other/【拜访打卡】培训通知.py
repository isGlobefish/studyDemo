# -*- coding: utf-8 -*-
'''
@createTool    : PyCharm-2020.3.2
@projectName   : pythonProjectPy3.9
@originalAuthor: Made in win10.Sys design by deHao.Zou
@createTime    : 2021/02/02 11:21
'''

import os
import re
import xlwt
import glob
import xlrd
import uuid
import time
import hmac
import json
import pyhdb
import base64
import shutil
import pymysql
import hashlib
import calendar
import requests
import datetime
import win32com
import pythoncom
import numpy as np
import urllib.parse
import pandas as pd
import urllib.request
from termcolor import cprint
from PIL import ImageGrab, Image
from time import strftime, gmtime
from qiniu import Auth, put_file, etag
from win32com.client import Dispatch, DispatchEx


# 定义钉钉功能
class dingdingFunction(object):
    def __init__(self, roboturl, robotsecret, appkey, appsecret):
        """
        :param roboturl: 群机器人WebHook_url
        :param robotsecret: 安全设置的加签秘钥
        :param appkey: 企业开发平台小程序AppKey
        :param appsecret: 企业开发平台小程序AppSecret
        """
        self.roboturl = roboturl
        self.robotsecret = robotsecret
        self.appkey = appkey
        self.appsecret = appsecret
        timestamp = round(time.time() * 1000)  # 时间戳
        secret_enc = robotsecret.encode('utf-8')
        string_to_sign = '{}\n{}'.format(timestamp, robotsecret)
        string_to_sign_enc = string_to_sign.encode('utf-8')
        hmac_code = hmac.new(secret_enc, string_to_sign_enc, digestmod=hashlib.sha256).digest()
        sign = urllib.parse.quote_plus(base64.b64encode(hmac_code))  # 最终签名
        self.webhook_url = self.roboturl + '&timestamp={}&sign={}'.format(timestamp, sign)  # 最终url,url+时间戳+签名

    # 发送文件
    def getAccess_token(self):
        url = 'https://oapi.dingtalk.com/gettoken?appkey=%s&appsecret=%s' % (AppKey, AppSecret)
        headers = {
            'Content-Type': "application/x-www-form-urlencoded"
        }
        data = {'appkey': self.appkey,
                'appsecret': self.appsecret}
        r = requests.request('GET', url, data=data, headers=headers)
        access_token = r.json()["access_token"]
        return access_token

    def getMedia_id(self, filespath):
        access_token = self.getAccess_token()  # 拿到接口凭证
        url = 'https://oapi.dingtalk.com/media/upload?access_token=' + access_token + '&type=file'
        files = {'media': open(filespath, 'rb')}
        data = {'access_token': access_token,
                'type': 'file'}
        response = requests.post(url, files=files, data=data)
        json = response.json()
        return json["media_id"]

    def sendFile(self, chatid, filespath):
        access_token = self.getAccess_token()
        media_id = self.getMedia_id(filespath)
        url = 'https://oapi.dingtalk.com/chat/send?access_token=' + access_token
        header = {
            'Content-Type': 'application/json'
        }
        data = {'access_token': access_token,
                'chatid': chatid,
                'msg': {
                    'msgtype': 'file',
                    'file': {'media_id': media_id}
                }}
        r = requests.request('POST', url, data=json.dumps(data), headers=header)
        print(r.json()["errmsg"])

    # 发送消息
    def sendMessage(self, content, chatName, num, sum):
        """
        :param content: 发送内容
        """
        header = {
            "Content-Type": "application/json",
            "Charset": "UTF-8"
        }
        sendContent = json.dumps(content)  # 将字典类型数据转化为json格式
        sendContent = sendContent.encode("utf-8")  # 编码为UTF-8格式
        request = urllib.request.Request(url=self.webhook_url, data=sendContent, headers=header)  # 发送请求
        opener = urllib.request.urlopen(request)  # 将请求发回的数据构建成为文件格式
        print('>>> ' + str(num).zfill(3) + '/' + str(sum).zfill(3) + ' ' + chatName)  # 返回发送结果


if __name__ == '__main__':
    orgChatData = pd.read_excel('D:/DataCenter/VisImage/伙伴对照表.xlsx', header=0)
    # orgChatData = pd.read_excel('C:/Users/Zeus/Desktop/伙伴对照表.xlsx', header=0)
    ChatData = orgChatData.drop_duplicates(['Partner']).reset_index(drop=True)

    AppKey = 'dingjpjkc2vaqjoqgmhz'  # 企业开发平台小程序AppKey
    AppSecret = 'oKNcuSF12oW0j9eBeO53wA6qwmKCVz34NVy1NvtvnjsvKPOdKiozsSZzUypNSWDc'  # 企业开发平台小程序AppSecret

    for ichat in range(len(ChatData)):
        try:

            ddMessage = {  # 发布消息内容
                "msgtype": "markdown",
                "markdown": {"title": "关于草晶华BI系统核查存疑内容收集的通知",  # @某人 才会显示标题
                             "text": "**【关于草晶华BI系统核查存疑内容收集的通知】**"
                                     "\n> ###### **各位草晶华的伙伴、同事们, 您们好！为助力破壁干出新规模, 达成目标。为保证我们数据的准确性, 现需要各位伙伴、各位同事登录各自的账号核查自己看板内容, 如若存在疑惑内容请反馈给数据运营中心邹德豪, 我们将全力为大家一一解决所有存在的问题, 感谢大家的支持。**"
                                     "\n> ###### **您可以通过以下两种方式登录：**"
                                     "\n> ###### **1、打开PC端钉钉, 点击工作台, 点击”PowerBI报表平台“, 直接进入(前提要有账号用户), 无需账号密码;**"
                                     "\n> ###### **2、打开浏览器（建议Chrome或者火狐浏览器）,[登录网址](http://pbi.zeus.cn/#/login):http://pbi.zeus.cn/#/login, 输入账号密码登录;**"
                                     "\n> ###### **恳请在使用过程中向我们提供宝贵的优化建议, 感谢您的大力支持, 谢谢。**"
                                     "\n###### ----------------------------------------------"
                                     "\n###### 发布时间：" + str(datetime.datetime.now()).split('.')[0]},  # 发布时间
                "at": {
                    # "atMobiles": [15817552982],  # 指定@某人
                    "isAtAll": True  # 是否@所有人[False:否, True:是]
                }
            }

            RobotWebHookURL = ChatData.loc[ichat, 'RobotURL']  # 群机器人url
            RobotSecret = ChatData.loc[ichat, 'RobotSecret']  # 群机器人加签秘钥secret(默认数运小跑腿)

            """
            特别说明：
                    发送消息：目前支持text、link、markdown等形式文字及图片，新增支持本地文件和图片类媒体文件的发送.
                    发送文件：目前支持简单excel表(csv、xlsx、xls等)、word、压缩文件,不支持ppt等文件的发送.
            """

            # 发送消息
            dingdingFunction(RobotWebHookURL, RobotSecret, AppKey, AppSecret).sendMessage(ddMessage, ChatData.loc[ichat, 'ChatName'], ichat + 1,
                                                                                          len(ChatData))
        except:
            pass
