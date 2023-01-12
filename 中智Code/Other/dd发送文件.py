# -*- coding: utf-8 -*-
'''
@createTool    : PyCharm-2020.2.2
@projectName   : pythonProject 
@originalAuthor: Made in win10.Sys design by deHao.Zou
@createTime    : 2020/11/11 11:38
'''
import requests
import json


def getAccess_token():
    appkey = 'ding7fxcfusxqkuv97pm'  # 管理员账号登录开发者平台，应用开发-创建应用-查看详情-appkey
    appsecret = 'FBEHIFyMG28nWZrn316df-ny5bmIz_LanRWtabCi'  # 应用里的appsecret
    url = 'https://oapi.dingtalk.com/gettoken?appkey=' + appkey + '&appsecret=' + appsecret
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
    path = 'C:/Users/Long/Desktop/nice.xlsx'  # 文件路径
    url = 'https://oapi.dingtalk.com/media/upload?access_token=' + access_token + '&type=file'
    files = {'media': open(path, 'rb')}
    data = {'access_token': access_token,
            'type': 'file'}
    response = requests.post(url, files=files, data=data)
    json = response.json()
    return json["media_id"]


def SendFile():
    access_token = getAccess_token()
    media_id = getMedia_id()
    chatid = 'DZnCErimkn2yQrn4aYel3JX7vPXKRonlvDFoVh1e'  # 通过jsapi工具获取的群聊id
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
    print(r.json())


if __name__ == '__main__':
    SendFile()
