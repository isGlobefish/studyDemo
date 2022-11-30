# -*- coding: utf-8 -*-
'''
@createTool    : PyCharm-2020.2.2
@projectName   : pythonProject 
@originalAuthor: Made in win10.Sys design by deHao.Zou
@createTime    : 2020/11/11 10:31
'''
import requests, json, re, io, time, hmac, hashlib, base64, random
from urllib.request import urlopen
from datetime import datetime
from PIL import Image, ImageFilter, ImageFont, ImageDraw
from bs4 import BeautifulSoup


# 高斯模糊类
class MyGaussianBlur(ImageFilter.Filter):
    name = "GaussianBlur"

    def __init__(self, radius=2, bounds=None):
        self.radius = radius
        self.bounds = bounds

    def filter(self, image):
        if self.bounds:
            clips = image.crop(self.bounds).gaussian_blur(self.radius)
            image.paste(clips, self.bounds)
            return image
        else:
            return image.gaussian_blur(self.radius)


# 卡片生成类
class myCard:
    def __init__(self):
        self.img_url = 'img/moban.jpg'  # 800*1280
        self.icon_url = 'img/icons.png'  # 200*400
        # 纵向间距
        self.magrinImgTop = 150
        # 横向间距
        self.magrinImgLeft = 30

        self.background_height = 0

    # 加载底图
    def loadMoban(self):
        self.img = Image.open(self.img_url)
        self.width, self.height = self.img.size

    # 加载图片
    def loadBackground(self, background_url):
        self.background = Image.open(background_url)

    # 加载小图
    def loadIcon(self):
        self.icon = Image.open(self.icon_url)

    # 高斯模糊图片作为背景
    def drawBlur(self):
        backflur = self.background.resize((self.width, self.height), resample=3).filter(MyGaussianBlur(radius=30))
        self.img.paste(backflur, (0, 0))

    # 添加标题
    def drawTitle(self):
        draw = ImageDraw.Draw(self.img)
        # 文字行距
        magrinTop = 20

        size = 150
        title_font = ImageFont.truetype('fzqk.TTF', size)
        x = self.width / 4 - size + self.magrinImgLeft
        y = self.magrinImgTop
        draw.text((x, y), '早', font=title_font, fill='#ffffff')

        size = 150
        title_font = ImageFont.truetype('fzqk.TTF', size)
        x = x
        y = y + size + magrinTop
        draw.text((x, y), '安', font=title_font, fill='#ffffff')

    # 添加图片
    def drawBackground(self):
        x = 0
        y = self.magrinImgTop * 2
        srcwidth, srcheight = self.background.size
        height = int(srcheight * self.width / srcwidth)
        # 重设图片尺寸
        background = self.background.resize((self.width, height), Image.ANTIALIAS)

        # 创建圆形遮罩
        alpha_layer = Image.new('L', (self.width, height), 0)
        draw = ImageDraw.Draw(alpha_layer)
        draw.ellipse((self.width / 2 - 100, 0, self.width - 30, height), fill=255)

        self.img.paste(background, (x, y), alpha_layer)
        self.background_height = y + height

    # 添加文字
    def drawText(self, tmp, weather, week):
        draw = ImageDraw.Draw(self.img)
        # 文字行距
        magrinTop = 180

        x = int(self.width / 2) - self.magrinImgLeft
        y = self.background_height + self.magrinImgTop
        font = ImageFont.truetype('wryh.ttf', 150)
        draw.text((x, y), tmp, font=font, fill='#ffffff')

        x = x
        y = y + magrinTop
        font = ImageFont.truetype('wryh.ttf', 50)
        draw.text((x, y), weather + ' ' + week, font=font, fill='#ffffff')

    # 添加小图
    def drawIcon(self):
        x_magrinImgLeft_add = 30
        x = self.magrinImgLeft + x_magrinImgLeft_add
        y = int(self.height / 2) + self.magrinImgTop
        self.img.paste(self.icon, (x, y), self.icon)

    # 保存到本地
    def saveCard(self):
        save_name = 'C:/Users/Long/Desktop/' + str(datetime.today()).split()[0] + '.jpg'
        self.img.save(save_name)
        return save_name

    def drawCard(self):
        self.loadMoban()

        background_url = getBingBackground()
        self.loadBackground(background_url)

        self.loadIcon()
        self.drawBlur()
        self.drawTitle()
        self.drawBackground()
        self.drawIcon()

        text = getWeatherNow('新城区,北京')
        if text != 0:
            self.drawText(text['tmp'], text['weather'], text['week'])
        else:
            print('Error:getWeatherNow() false')


# 卡片发送类
class sendCard:
    def __init__(self, save_name):
        self.whtoken = "https://oapi.dingtalk.com/robot/send?access_token=c1f27f62eeb27b6542c86e4009c5a69d3b8f7de4e9653fb3cf356686422554dc"

        self.getMediaid(save_name)
        self.text = getOne()

    # 上传图片得到资源id
    def getMediaid(self, img):
        APPKEY = 'DZnCErimkn2yQrn4aYel3JX7vPXKRonlvDFoVh1e'
        APPSECRET = 'FBEHIFyMG28nWZrn316df-ny5bmIz_LanRWtabCi'
        ACCESS_TOKEN = json.loads(
            requests.get('https://oapi.dingtalk.com/gettoken?appkey=' + APPKEY + '&appsecret=' + APPSECRET).text).get(
            'access_token')

        url = 'https://oapi.dingtalk.com/media/upload?access_token=' + ACCESS_TOKEN + '&type=image'
        files = {'media': open(img, 'rb')}
        data = {'access_token': ACCESS_TOKEN, 'type': 'image'}
        response = requests.post(url, files=files, data=data)
        if response.status_code == 200:
            img_res = response.json()
            self.img_id = img_res["media_id"]
        else:
            self.img_id = ''
            print('Error:getMediaid() false:' + str(response.status_code))

    # 发送Markdown
    def sendImg(self):
        timestamp = str(round(time.time() * 1000))
        app_secret = 'SECf1ff53603573024a6ac2804923b008ef865f432f74346b185336a89e5902515b'
        app_secret_enc = app_secret.encode('utf-8')
        string_to_sign = '{}\n{}'.format(timestamp, app_secret)
        string_to_sign_enc = string_to_sign.encode('utf-8')
        hmac_code = hmac.new(app_secret_enc, string_to_sign_enc, digestmod=hashlib.sha256).digest()
        sign = base64.b64encode(hmac_code).decode('utf-8')

        webhook = "https://oapi.dingtalk.com/robot/send?access_token=" + self.whtoken + "&timestamp=" + timestamp + "&sign=" + sign
        header = {"Content-Type": "application/json", "Charset": "UTF-8"}
        message = {
            "msgtype": "markdown",
            "markdown": {
                "title": '每日推送',
                "text": '> ![screenshot](%s) \n\n> %s' % (self.img_id, self.text)
            },
            "at": {
                "atDingtalkIds": [],
                "isAtAll": False
            }
        }
        message_json = json.dumps(message)
        info = requests.post(url=webhook, data=message_json, headers=header)
        print(info.text)


# 获取和风实时天气
def getWeatherNow(address='朝阳区,北京'):
    key = '和风天气API的key'
    now_url = 'https://free-api.heweather.net/s6/weather/now?key=%s&location=%s' % (key, address)
    now_info = requests.get(now_url)
    if (now_info.status_code == 200):
        now_result = json.loads(now_info.text).get('HeWeather6')[0]
        if (now_result.get('status') == 'ok'):
            weather_time = now_result.get('update').get('loc')
            weather = now_result.get('now')

            # 获取当前温度
            tmp = weather.get('tmp') + '℃'
            # 获取当前天气
            weather = weather.get('cond_txt')
            # 根据日期判断是星期几
            week = datetime.strptime(weather_time.split()[0], "%Y-%m-%d").weekday()
            weekdict = {0: '星期一', 1: '星期二', 2: '星期三', 3: '星期四', 4: '星期五', 5: '星期六', 6: '星期天', }

            return {'tmp': tmp, 'weather': weather, 'week': weekdict[week]}
        else:
            print('now_result.status is not ok')
            return 0
    else:
        print('now_info.status_code is not 200')
        return 0


# 获取bing每日壁纸
def getBingBackground():
    url = 'https://www.bing.com/'
    html = requests.get(url).text
    Nurl = re.findall('id="bgLink" rel="preload" href="(.*?)&', html, re.S)
    # 获取图片地址
    for temp in Nurl:
        url = 'https://www.bing.com' + temp
    # 图片不保存直接转换为流
    image_bytes = urlopen(url).read()
    data_stream = io.BytesIO(image_bytes)
    return data_stream


# 获取one每日一句
def getOneData(url):
    response = requests.get(url)
    if response.status_code != 200:
        return 0
    res = BeautifulSoup(response.text, "html.parser")
    for meta in res.select('meta'):
        if meta.get('name') == 'description':
            # 获取文字内容
            text = meta.get('content')
    return text


def getOne():
    one_content = 0
    count = 0
    # 范围内随机搜索一个页面，超过50次没有内容则获取失败
    while (one_content == 0):
        one_content = getOneData('http://wufazhuce.com/one/' + str(random.randint(14, 2850)))
        count = count + 1
        if count > 50:
            return ''
    return one_content


if __name__ == '__main__':
    # 生成卡片,在./img进行操作
    # card = myCard()
    # card.drawCard()
    # save_name = card.saveCard()
    # 发送卡片
    pot = sendCard('C:/Users/Long/Desktop/12.jpg')
    pot.sendImg()
