# -*- coding: utf-8 -*-
'''
@createTool    : PyCharm-2020.2.2
@projectName   : pythonCode 
@originalAuthor: Made in win10.Sys design by deHao.Zou
@createTime    : 2020/10/12 11:37
'''

import re
import time
import random
import requests
import json
import pandas as pd

name, specs, market, date, price, website = [], [], [], [], [], []
A = ['11306990003220200', '13416990003220200', '15101990003220200', '14509990003220200', '11306990004780200',
     '13416990004780200', '15101990004780200', '14509990004780200', '11306990006090700', '15101990006090700',
     '14509990006090700', '14509990006880300', '14509990001740300 ', '11306990001740300 ', '15106990001740300 ',
     '15101990003100400']  # 天地编号
A1 = ['安国药市', '亳州药市', '荷花池药市', '玉林药市', '安国药市', '亳州药市', '荷花池药市', '玉林药市', '安国药市', '荷花池药市', '玉林药市', '玉林药市', '玉林药市', '安国药市',
      '荷花池药市', '荷花池药市']  # 市场
A2 = ['菊花', '菊花', '菊花', '菊花', '肉苁蓉', '肉苁蓉', '肉苁蓉', '肉苁蓉', '西洋参', '西洋参', '西洋参', '鱼腥草', '甘草', '甘草', '甘草', '金银花']  # 品名
A3 = ['杭菊胎 浙江', '杭菊胎 浙江', '杭菊胎 浙江', '杭菊胎 浙江', '硬个 新疆', '硬个 新疆', '硬个 新疆', '硬个 新疆', '长支 国产', '长支 国产', '长支 国产', '家统 广西',
      '统片 甘肃', '统片 甘肃', '统片 甘肃', '色白花全花蕾 山东']  # 规格

B1 = ['三七', '三七', '三七', '三七', '白芍', '白芍', '白芍', '丹参', '丹参', '当归', '当归', '当归', '当归', '茯苓', '茯苓', '茯苓', '茯苓', '红参', '红景天',
      '红景天', '红景天', '红景天', '黄芪', '黄芪', '黄芪', '黄芪', '决明子', '罗布麻叶', '罗布麻叶', '罗布麻叶', '罗汉果', '罗汉果', '罗汉果', '罗汉果', '玫瑰花',
      '玫瑰花', '山楂', '山楂', '山楂', '山楂', '天麻', '淫羊藿', '淫羊藿', '桔梗', '桔梗', '桔梗', '桔梗', '苦杏仁', '连翘', '连翘', '连翘']  # 品名
B2 = ['60头', '60头', '60头', '60头', '统', '统', '统', '精品', '精品', '散把', '散把', '散把', '散把', '块', '块', '块', '块', '30支无糖', '大花',
      '大花', '大花', '大花', '中条', '中条', '中条', '中条', '包含量', '统', '统', '统', '1200个/箱', '1200个/箱', '1200个/箱', '1200个/箱', '炕货',
      '炕货', '手工片', '手工片', '手工片', '手工片', '小', '优', '优质', '统', '统', '统', '统', '统', '青水煮', '青水煮', '青水煮']  # 规格
B3 = ['云南', '云南', '云南', '云南', '安徽', '安徽', '安徽', '山东', '山东', '甘肃', '甘肃', '甘肃', '甘肃', '安徽', '安徽', '安徽', '安徽', '东北', '西藏',
      '西藏', '西藏', '西藏', '甘肃', '甘肃', '甘肃', '甘肃', '国产', '天津', '天津', '天津', '广西', '广西', '广西', '广西', '甘肃', '甘肃', '山东', '山东',
      '山东', '山东', '安徽', '甘肃', '甘肃', '安徽', '安徽', '安徽', '安徽', '内蒙', '山西', '山西', '山西']  # 产地
B4 = ['亳州市场', '安国市场', '成都市场', '玉林市场', '亳州市场', '成都市场', '玉林市场', '亳州市场', '成都市场', '亳州市场', '安国市场', '成都市场', '玉林市场', '亳州市场',
      '安国市场', '成都市场', '玉林市场', '亳州市场', '亳州市场', '安国市场', '成都市场', '玉林市场', '亳州市场', '安国市场', '成都市场', '玉林市场', '亳州市场', '亳州市场',
      '安国市场', '玉林市场', '亳州市场', '安国市场', '成都市场', '玉林市场', '亳州市场', '安国市场', '亳州市场', '安国市场', '成都市场', '玉林市场', '亳州市场', '亳州市场',
      '安国市场', '成都市场', '玉林市场', '亳州市场', '安国市场', '亳州市场', '亳州市场', '安国市场', '成都市场']  # 市场

url = 'https://www.zyctd.com/Breeds/GetPriceChart'
headers = {
    "Host": "www.zyctd.com",
    "User-Agent": "Mozilla/5.0 (Windows NT 6.1; Win64; x64; rv:54.0) Gecko/20100101 Firefox/54.0",
    "Accept": "*/*",
    "Accept-Language": "zh-CN,zh;q=0.8,en-US;q=0.5,en;q=0.3",
    "Accept-Encoding": "gzip, deflate, br",
    "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
    "X-Requested-With": "XMLHttpRequest",
    "Referer": "https://www.zyctd.com/jiage/xq70.html",
    "Content-Length": "35",
    "Cookie": "FromsAuthByDbCookie_zytd_Edwin.PrvGuest=1c1uuuuuuuuuuaIab597WcT17c4818Vbc1SaSc1U5acUWcbU4WS0o5qqnc78ca99o698qac9nadmoll5moad015; UM_distinctid=171722e10003d-0a2865dac79651-17397540-15f900-171722e1001341; CNZZDATA1261355531=1147626159-1586755002-%7C1586851405; Hm_lvt_ba57c22d7489f31017e84ef9304f89ec=1586758554,1586830545,1586852111; Hm_lpvt_ba57c22d7489f31017e84ef9304f89ec=1586852111",
    "Connection": "keep-alive"
}
for j in range(len(A)):
    try:
        # print("获取页数:",j)
        stop = random.uniform(1, 3)
        data = {"PriceType": "day", "mid": A[j]}
        response = requests.post(url=url, data=data, headers=headers, timeout=10)
        content0 = json.loads(response.text)
        content = content0['Data']['PriceChartData']
        it = re.findall('\\[(.*?),(.*?)\\]', content)
        for i in range(len(it)):
            # list2 = []
            name.append(A2[j])
            specs.append(A3[j])
            market.append(re.sub('药市', '', A1[j]))
            if i == 0:
                a1 = it[i][0].replace(it[0][0][0], '')
            else:
                a1 = it[i][0]
            timeStamp = int(a1) / 1000
            timeArray = time.localtime(timeStamp)
            otherStyleTime = time.strftime("%Y-%m-%d", timeArray)
            date.append(otherStyleTime)
            price.append(it[i][1])
            website.append("天")
            # list1.append(list2)

    except:
        print("出错", j)
url = 'https://www.yt1998.com/price/historyPriceQ!getHistoryPrice.do'

headers = {
    "Host": "www.yt1998.com",
    "Connection": "keep-alive",
    "Content-Length": "42",
    "Accept": "application/json, text/javascript, */*; q=0.01",
    "Origin": "https://www.yt1998.com",
    "X-Requested-With": "XMLHttpRequest",
    "User-Agent": "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/75.0.3770.100 Safari/537.36",
    "Content-Type": "application/x-www-form-urlencoded;charset=UTF-8",
    "Referer": "https://www.yt1998.com/priceHistory.html?keywords=%E7%9F%B3%E5%88%81%E6%9F%8F&guige=%E7%BB%9F&chandi=%E6%96%B0%E7%96%86&market=1",
    "Accept-Encoding": "gzip, deflate, br",
    "Accept-Language": "zh-CN,zh;q=0.9",
    "Cookie": "JSESSIONID=ADC6AAA4C1429B382CCA1BAABB8E688D; Hm_lvt_21f2fde8228a3428719fdc5669ab5410=1586739170,1586831855,1587015027,1588055351; Hm_lpvt_21f2fde8228a3428719fdc5669ab5410=1588055496"

}
for j in range(len(B1)):
    try:
        # print("获取页数:",j)
        stop = random.uniform(1, 3)
        data = {"ycnam": B1[j], "guige": B2[j], "chandi": B3[j], "market": B4[j]}
        response = requests.post(url=url, data=data, headers=headers, timeout=10)
        content0 = json.loads(response.text)
        content = content0['data']

        for i in range(len(content)):
            list2 = []
            name.append(B1[j])
            specs.append(B2[j] + ' ' + B3[j])
            market.append(re.sub('市场', '', B4[j]))
            date.append(content[i]['Date_time'])
            price.append(content[i]['DayCapilization'])
            website.append("通")
    except:
        print("出错", j)
Attractions = pd.DataFrame({'品名': name, '规格': specs, '市场': market, '日期': date, '价格': price, '网站': website})
# print(Attractions)
