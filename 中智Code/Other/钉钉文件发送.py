# -*- coding: utf-8 -*-
'''
@createTool    : PyCharm-2020.2.2
@projectName   : pythonProject 
@originalAuthor: Made in win10.Sys design by deHao.Zou
@createTime    : 2020/11/11 11:01
'''
import dingtalk.api
import requests

class DingDingAPI():
    def __init__(self, appkey, appsecret):
        self.appkey = appkey
        self.appsecret = appsecret
        self.access_token = self.get_token()

    # 获取token
    def get_token(self):
        params = {
            "appkey": self.appkey,
            "appsecret": self.appsecret
        }
        try:
            res = requests.get("https://oapi.dingtalk.com/gettoken", params=params)
            access_token = res.json().get("access_token")
            return access_token
        except Exception as e:
            print(e)

    # 发送消息
    def send_msg(self, chatid="", content=""):

        req = dingtalk.api.OapiChatSendRequest("https://oapi.dingtalk.com/chat/send")

        req.chatid = chatid
        req.text = {
            "content": content
        }
        req.msgtype = "text"

        resp = req.getResponse(self.access_token)

        return resp

    # 发送文件
    def send_file(self, chatid="", media_id=""):

        req = dingtalk.api.OapiChatSendRequest("https://oapi.dingtalk.com/chat/send")

        req.chatid = chatid
        req.file = {
            "media_id": media_id
        }
        req.msgtype = "file"
        resp = req.getResponse(self.access_token)
        return resp


    # 上传文件
    def upload_media(self,file_name="name.docx",file_path=""):
        req = dingtalk.api.OapiMediaUploadRequest("https://oapi.dingtalk.com/media/upload")

        req.type = "file"
        req.media = dingtalk.api.FileItem(file_name,open(file_path, 'rb'))
        resp = req.getResponse(self.access_token)
        return resp

    # 通过电话获取userid
    def get_by_mobile(self,mobile):
        req = dingtalk.api.OapiUserGetByMobileRequest("https://oapi.dingtalk.com/user/get_by_mobile")

        req.mobile = mobile
        try:
            resp = req.getResponse(self.access_token)
            return resp
        except Exception as e:
            print(e)

    # 通过userid获取信息,不能使用机器人appkey
    def get_user_msg(self,userid):
        req = dingtalk.api.OapiUserGetRequest("https://oapi.dingtalk.com/user/get")
        req.userid = userid
        try:
            resp = req.getResponse(self.access_token)
            print(resp)
            return resp
        except Exception as e:
            print(e)


if __name__ == '__main__':
    appkey = "ding7fxcfusxqkuv97pm"
    appsecret = "xW8ILzIcg1w_I6Rz_B1PQJxFOD0e3TYEYaPtXmHt-9GdLyC_gt-NVKO9ciRyKwAH"
    dd=DingDingAPI(appkey,appsecret)
    dd.send_msg()