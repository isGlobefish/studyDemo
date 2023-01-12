# -*- coding: utf-8 -*-
'''
@createTool    : PyCharm-2020.2.2
@projectName   : pythonProject 
@originalAuthor: Made in win10.Sys design by deHao.Zou
@createTime    : 2020/11/9 10:22
'''
import os
import glob
import datetime
import openpyxl
import pandas as pd
from termcolor import cprint
from xlrd import xldate_as_tuple
from time import strftime, gmtime
from sqlalchemy import create_engine  # 连接mysql使用
from sqlalchemy.types import Integer, NVARCHAR, Float

df = pd.read_excel("C:/Users/Long/Desktop/nice.xlsx", sheet_name="老百姓11月", dtype=str)
df["地区"] = ''
df["过账日期"] = ''

from xlwt.Workbook import *
from xlwt.Style import *
from xlrd import open_workbook
from xlutils.copy import copy
import xlrd

style = XFStyle()
rb = open_workbook("C:/Users/Long/Desktop/nice.xlsx", formatting_info=True)
wb = copy(rb.get_sheet(1))

new_book = Workbook()
w_sheet = wb.get_sheet(0)
w_sheet.write(6, 6)

wb.save("C:/Users/Long/Desktop/copy123.xlsx")

import pandas as pd

df1 = pd.DataFrame({'a': [3, 1], 'b': [4, 3]})
df2 = df1.copy()
with pd.ExcelWriter('C:/Users/Long/Desktop/output.xlsx') as writer:
    str1 = ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q']
    for i in str1:
        name = str(i)
        df1.to_excel(writer, sheet_name=name)
writer.save()
writer.close()

from dingtalkchatbot.chatbot import DingtalkChatbot, FeedLink
import time
import hmac
import urllib
import hashlib
import base64


# 获取链接,填入urlToken 和 secret
def getSIGN():
    timestamp = str(round(time.time() * 1000))
    urlToken = "https://oapi.dingtalk.com/robot/send?access_token=c1f27f62eeb27b6542c86e4009c5a69d3b8f7de4e9653fb3cf356686422554dc"
    secret = 'SECf1ff53603573024a6ac2804923b008ef865f432f74346b185336a89e5902515b'
    secret_enc = secret.encode('utf-8')
    string_to_sign = '{}\n{}'.format(timestamp, secret)
    string_to_sign_enc = string_to_sign.encode('utf-8')
    hmac_code = hmac.new(secret_enc, string_to_sign_enc, digestmod=hashlib.sha256).digest()
    sign = urllib.parse.quote_plus(base64.b64encode(hmac_code))
    SignMessage = urlToken + "&timestamp=" + timestamp + "&sign=" + sign
    return SignMessage


SignMessage = getSIGN()
xiaoDing = DingtalkChatbot(SignMessage)  # 初始化机器人


def toSend():
    """第一: 发送文本-->
    send_text(self,msg,is_at_all=False,at_mobiles=[],at_dingtalk_ids=[],is_auto_at=True)
        msg: 发送的消息
        is_at_all:是@所有人吗? 默认False,如果是True.会覆盖其它的属性
        at_mobiles:要@的人的列表,填写的是手机号
        at_dingtalk_ids:未知;文档说的是"被@人的dingtalkId（可选）"
        is_auto_at:默认为True.经过测试,False是每个人一条只能@一次,重复的会过滤,否则不然,测试结果与文档不一致
    """
    xiaoDing.send_text("自动就好 我是海王深圳鸭蛋1号,我为自己代言", is_at_all=False)


toSend()


def sentPicture():
    """第二:发送图片
    send_image(self, pic_url):
        pic_url: "图片地址"
    """
    xiaoDing.send_image("http://rrd.me/gE93L")


sentPicture()


def 发送link():
    """第三:发送link
    send_link(self, title, text, message_url, pic_url='')
        title:标题    text:内容,太长会自动截取
        message_url:跳转的url  pic_url:添加的图片的url(可选)
    """
    xiaoDing.send_link(title="今天是星期8", text="牵你的手，朝朝暮暮，牵你的手，等待明天，牵你的手，走过今生，牵你的手，生生世世",
                       message_url="https://baidu.com", pic_url="http://rrd.me/gE93L")


def markdown():
    """第四:发送markdown
    send_markdown(self,title,text,is_at_all=False,at_mobiles=[],at_dingtalk_ids=[],is_auto_at=True)
        title:标题    text:内容
        is_at_all: @所有人时：true，否则为：false（可选）
        at_mobiles: 被@人的手机号（默认自动添加在text内容末尾，可取消自动化添加改为自定义设置，可选）
        at_dingtalk_ids: 被@人的dingtalkId（可选）
        is_auto_at: 是否自动在text内容末尾添加@手机号，默认自动添加，可设置为False取消（可选）
    """
    xiaoDing.send_markdown(title="我是标题", text="我是内容,啊哈哈哈哈哈", is_at_all=True)


markdown()


def catLink():
    # send_feed_card(links)
    """
        links是一个列表a,列表里每个元素又是列表b
        列表b的属性:
            title:标题    message_url:点开后跳转的URL   pic_url:图片的地址
    Returns:

    """
    feedlink1 = FeedLink(title="猫1", message_url="https://www.badiu.com/",
                         pic_url="http://rrd.me/gE9zB")
    feedlink2 = FeedLink(title="猫2", message_url="https://www.badiu.com/",
                         pic_url="http://rrd.me/gE9zN")
    feedlink3 = FeedLink(title="猫3", message_url="https://www.badiu.com/",
                         pic_url="http://rrd.me/gE9zV")
    feedlin4k = FeedLink(title="猫4", message_url="https://www.badiu.com/",
                         pic_url="http://rrd.me/gE92a")

    links = [feedlink1, feedlink2, feedlink3, feedlin4k]
    xiaoDing.send_feed_card(links)


catLink()

import time
import hmac
import hashlib
import base64
import json
import requests
import urllib.parse
import urllib.request


class dingdingFunction(object):
    def __init__(self, secret=None, url=None):
        """
        :param url: 机器人没有加签的WebHook_url
        :param secret: 安全设置的加签秘钥
        """
        if url is not None:
            url = url
        else:
            url = 'https://oapi.dingtalk.com/robot/send?access_token=a0ebec1fda6cb1dcfbb6777d90ac61dfe097746b277358698593a473156b361f'  # 无加密的url
        if secret is not None:
            secret = secret
        else:
            secret = 'xW8ILzIcg1w_I6Rz_B1PQJxFOD0e3TYEYaPtXmHt-9GdLyC_gt-NVKO9ciRyKwAH'  # 加签秘钥
        timestamp = round(time.time() * 1000)  # 时间戳
        secret_enc = secret.encode('utf-8')
        string_to_sign = '{}\n{}'.format(timestamp, secret)
        string_to_sign_enc = string_to_sign.encode('utf-8')
        hmac_code = hmac.new(secret_enc, string_to_sign_enc, digestmod=hashlib.sha256).digest()
        sign = urllib.parse.quote_plus(base64.b64encode(hmac_code))  # 最终签名
        self.webhook_url = url + '&timestamp={}&sign={}'.format(timestamp, sign)  # 最终url，url+时间戳+签名

    def sendMessage(self, content):
        """
        发送消息至机器人对应的群
        :param content: 发送的内容
        :return:
        """
        header = {
            "Content-Type": "application/json",
            "Charset": "UTF-8"
        }
        sendContent = json.dumps(content)  # 将字典类型数据转化为json格式
        sendContent = sendContent.encode("utf-8")  # 编码为UTF-8格式
        request = urllib.request.Request(url=self.webhook_url, data=sendContent, headers=header)  # 发送请求
        opener = urllib.request.urlopen(request)  # 将请求发回的数据构建成为文件格式
        print(opener.read())  # 打印返回的结果


#     加一个发送文件（包括图片、文本、表格、压缩文件等等）
#     获取全部手机号（匹配人名与手机号）


if __name__ == '__main__':
    ddWebHookURL = 'https://oapi.dingtalk.com/robot/send?access_token=a0ebec1fda6cb1dcfbb6777d90ac61dfe097746b277358698593a473156b361f'
    ddSecret = 'xW8ILzIcg1w_I6Rz_B1PQJxFOD0e3TYEYaPtXmHt-9GdLyC_gt-NVKO9ciRyKwAH'
    ddMessage = {
        "msgtype": "markdown",
        "markdown": {"title": "测试钉钉机器人ing",
                     "text": "# 一级标题 \n## 二级标题 \n> 引用文本  \n**加粗**  \n*斜体*  \n[百度链接](https://www.baidu.com) "
                     "\n![草莓](https://dss0.bdstatic.com/70cFuHSh_Q1YnxGkpoWK1HF6hhy/it/u=1906469856,4113625838&fm=26&gp=0.jpg) "
                             "\n![screenshot](C:/Users/Long/Desktop/4.png) "
                             "\n- 无序列表 \n1.有序列表  \n@某手机号主 @+86-13267854059"},
        "at": {
            "atMobiles": [13267854059],
            "isAtAll": False  # 是否@所有人
        }
    }
    dingding = dingdingFunction(secret=ddSecret, url=ddWebHookURL)
    dingding.sendMessage(ddMessage)







import time
import hmac
import hashlib
import base64
import json
import requests
import urllib.parse
import urllib.request


def getAccess_token():
    appkey = 'dingbckcbv5kccupsp4s'
    appsecret = 'yIf8LqE66Gcbib3jm8KyuGoVt1-NdTQoj57HPi81TFJc8cm01PGPZHDGC6VJpirv'

    url = 'https://oapi.dingtalk.com/gettoken?appkey=%s&appsecret=%s' % (appkey, appsecret)

    headers = {
        'Content-Type': "application/x-www-form-urlencoded"
    }
    data = {'appkey': appkey,
            'appsecret': appsecret}
    r = requests.request('GET', url, data=data, headers=headers)
    access_token = r.json()["access_token"]
    return access_token


def getMedia_id():
    access_token = getAccess_token()  # 拿到接口凭证
    filesPath = 'C:/Users/Long/Desktop/2020.11.01-10海王深圳.xlsx'  # 文件路径
    url = 'https://oapi.dingtalk.com/media/upload?access_token=' + access_token + '&type=file'
    files = {'media': open(filesPath, 'rb')}
    data = {'access_token': access_token,
            'type': 'file'}
    response = requests.post(url, files=files, data=data)
    json = response.json()
    return json["media_id"]


def sendTodingding():
    access_token = getAccess_token()
    media_id = getMedia_id()
    chatid = 'chat913e2469633a29f5cb7abe1b4970f3d1'  # 通过jsapi工具获取的群聊id
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


sendTodingding()