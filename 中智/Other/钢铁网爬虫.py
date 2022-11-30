# -*- coding: utf-8 -*-
"""
Created on Tue Apr 14 14:42:38 2020

@author: Administrator
"""
import re
import requests
import json
import importlib, sys

importlib.reload(sys)
import pandas as pd

headers = {
    "Host": "mysteelapi.steelphone.com",
    "Connection": "keep-alive",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/81.0.4044.138 Safari/537.36",
    "Accept": "*/*",
    "Sec-Fetch-Site": "cross-site",
    "Sec-Fetch-Mode": "no-cors",
    "Sec-Fetch-Dest": "script",
    "Referer": "https://duxiban.mysteel.com/",
    "Accept-Encoding": "gzip, deflate, br",
    "Accept-Language": "zh-CN,zh;q=0.9"
}

urls = []
urls.append(
    'https://mysteelapi.steelphone.com/tpl/zhanting_data.html?indexCodes=ST_0000099855,ST_0000108449,ST_0000096209&startTime=2019-5-11&callback=jQuery18309638561318193197_1589179366429&_=1589179870464')
urls.append(
    'https://mysteelapi.steelphone.com/tpl/zhanting_data.html?indexCodes=ST_0000074400,ST_0000075148,ST_0000073973&startTime=2019-5-11&callback=jQuery18309638561318193197_1589179366429&_=1589183292174')
urls.append(
    'https://mysteelapi.steelphone.com/tpl/zhanting_data.html?indexCodes=YS_0000065990&startTime=2019-5-11&callback=jQuery18309638561318193197_1589179366429&_=1589184026781')
name = {'镀锡': ['天津', '上海', '广州'], '热轧4.75mm': ['上海', '广州', '天津'], '泸锡': ['上海']}
name1 = ['镀锡', '热轧4.75mm', '泸锡']
list1 = []
tiename, city, price, date = [], [], [], []
for k in range(len(urls)):
    response = requests.get(url=urls[k], headers=headers, timeout=10)
    response1 = re.findall('\\((.*)\\)', response.text)
    content0 = json.loads(response1[0])
    content4 = content0['xAxis']  # 日期
    try:
        for i in range(len(name[name1[k]])):
            for j in range(len(content4)):
                list2 = []
                tiename.append(name1[k])
                city.append(name[name1[k]][i])
                price.append(content0['datas'][i]['yAxis'][j])
                date.append(content0['xAxis'][j])
    except:
        print("出错", name1[k] + name[name1[k]][i])
tietable = pd.DataFrame({'name': tiename, 'city': city, 'price': price, 'date': date})
tietable.to_excel('C:/Users/Long/Desktop/123.xlsx', index=False)
# print(tietable)
