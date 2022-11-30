# -*- coding: utf-8 -*-
'''
@createTool    : PyCharm-2020.2.2
@projectName   : pythonProjectPy3.9 
@originalAuthor: Made in win10.Sys design by deHao.Zou
@createTime    : 2020/12/3 17:53
'''
# import pandas as pd
#
# url = 'http://www.kuaidaili.com/free/'
# df = pd.read_html(url)[0]
# # [0]：表示第一个table，多个table需要指定，如果不指定默认第一个
# # 如果没有【0】，输入dataframe格式组成的list
# df
# print(type(df))
# # df.to_csv('C:/Users/Long/Desktop/free_ip.csv', mode='a', encoding='utf_8_sig', header=0, index=0)
# df.to_excel('C:/Users/Long/Desktop/free_ip.xlsx', header=0, index=0, encoding='utf-8')
# print('done!')

# 抓取 ajax 页面
# demo 3

import urllib.request, urllib.parse, urllib.error
import urllib.parse
import urllib.error
import json

url = 'https://movie.douban.com/j/new_search_subjects?'
headers = {
    'Accept': 'text/html',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/68.0.3440.106 Safari/537.36',
}

# 循环次数自己设置，这里循环三次，爬取前60个电影数据
for p in range(0, 3):
    page = p * 20
    params = {'sort': 'U', 'range': '0,10', 'tags': '', 'start': str(page)}
    params_encode = urllib.parse.urlencode(params).encode('utf-8')
    try:
        request = urllib.request.Request(url, data=params_encode, headers=headers)
        with urllib.request.urlopen(request) as response:
            print(json.loads(response.read().decode('utf-8')))
            print('---' * 20)
    except urllib.error.HTTPError as e:
        print(e)
    except urllib.error.URLError as e:
        print(e)
