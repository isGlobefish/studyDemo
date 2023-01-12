# -*- coding: utf-8 -*-
'''
@createTool    : PyCharm-2020.2.2
@projectName   : pythonProject 
@originalAuthor: Made in win10.Sys design by deHao.Zou
@createTime    : 2020/11/12 10:51
'''
import time
import hmac
import json
import base64
import hashlib
import requests
import urllib.parse
import urllib.request


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
        self.webhook_url = self.roboturl + '&timestamp={}&sign={}'.format(timestamp, sign)  # 最终url，url+时间戳+签名

    # 发送消息
    def sendMessage(self, content):
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
        print(opener.read().decode())  # 打印返回的结果

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


#     加一个发送文件（包括图片、文本、表格、压缩文件等等）
#     获取全部手机号（匹配人名与手机号）


if __name__ == '__main__':
    FilePath = 'C:/Users/Long/Desktop/4.png'  # 待发送文件路径
    ChatId = 'chat294fd3795ede63fc6a479e5c074f9ba2'  # 通过jsapi工具https://wsdebug.dingtalk.com/?spm=a219a.7629140.0.0.7bc84a972WUfGd获取目标群聊id
    AppKey = 'dingjpjkc2vaqjoqgmhz'  # 企业开发平台小程序AppKey
    AppSecret = 'oKNcuSF12oW0j9eBeO53wA6qwmKCVz34NVy1NvtvnjsvKPOdKiozsSZzUypNSWDc'  # 企业开发平台小程序AppSecret
    RobotWebHookURL = 'https://oapi.dingtalk.com/robot/send?access_token=dd024c8278110ff67cc706c1cc44234b3469f2e44fb9b5e1c17eecae713ad94c'  # 群机器人url
    RobotSecret = 'GbSFeeIHgYNJfXT5WoPT6c6GRmMVRd2wVODyexo7SQIF5HJkucowab6cNMiyR8IV'  # 群机器人加签秘钥secret
    ddMessage = {  # 消息内容
        "msgtype": "markdown",
        "markdown": {"title": "测试钉钉机器人ing",
                     "text": "# 一级标题 \n## 二级标题 \n> 引用文本  \n**加粗**  \n*斜体*  \n[百度链接](https://www.baidu.com) "
                             "\n![草莓](https://dss0.bdstatic.com/70cFuHSh_Q1YnxGkpoWK1HF6hhy/it/u=1906469856,4113625838&fm=26&gp=0.jpg) "
                             "\n- 无序列表 \n1.有序列表  \n@某手机号主 @13267854059"},
        "at": {
            "atMobiles": [13267854059],
            "isAtAll": False  # 是否@所有人
        }
    }
    """
    特别说明：
            发送消息：目前支持text、link、markdown等形式文字及图片，并不支持本地文件和图片类媒体文件的发送
            发送文件：目前支持简单excel表(csv、xlsx、xls等)、word、压缩文件，不支持ppt等文件的发送
    """
    dingdingFunction(RobotWebHookURL, RobotSecret, AppKey, AppSecret).sendFile(ChatId, FilePath)  # 发送文件
    dingdingFunction(RobotWebHookURL, RobotSecret, AppKey, AppSecret).sendMessage(ddMessage)  # 发送消息
