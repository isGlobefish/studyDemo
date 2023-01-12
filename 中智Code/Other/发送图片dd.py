import requests
import json


class Dingding_Robot_new():
    def __init__(self, dingding_cfg=None):
        self.appkey = dingding_cfg.appkey
        self.appsecret = dingding_cfg.appsecret
        self.chatid = dingding_cfg.chatid

    def getAccess_token(self):
        appkey = self.appkey  # 管理员账号登录开发者平台，应用开发-创建应用-查看详情-appkey
        appsecret = self.appsecret  # 应用里的appsecret
        # https://oapi.dingtalk.com/gettoken?appkey=key&appsecret=secret
        url = 'https://oapi.dingtalk.com/gettoken'
        headers = {
            'Content-Type': "application/x-www-form-urlencoded"
        }
        data = {'appkey': appkey,
                'appsecret': appsecret}
        r = requests.get(url=url, params=data, headers=headers)
        # print(r.text)
        access_token = r.json()["access_token"]

        print(access_token)
        return access_token

    def getMedia_id(self):
        access_token = self.getAccess_token()  # 拿到接口凭证
        path = 'C:/Users/Long/Desktop/123.txt'  # 文件路径
        url = 'https://oapi.dingtalk.com/media/upload?access_token=' + access_token + '&type=file'
        files = {'media': open(path, 'rb')}
        data = {'access_token': access_token,
                'type': 'file'}
        response = requests.post(url, files=files, data=data)
        json = response.json()
        print(json)
        return json["media_id"]

    def SendFile(self):
        access_token = self.getAccess_token()
        media_id = self.getMedia_id()
        chatid = self.chatid  # 通过jsapi工具获取的群聊id
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


if __name__ == "__main__":
    run = Dingding_Robot_new()
    run.SendFile()
