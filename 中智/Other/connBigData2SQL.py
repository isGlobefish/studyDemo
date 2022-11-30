# -*- coding: utf-8 -*-
'''
@createTool    : PyCharm-2020.2.2
@projectName   : pythonProject 
@originalAuthor: Made in win10.Sys design by deHao.Zou
@createTime    : 2020/10/21 9:34
'''

import pymysql
import pandas as pd

conn = pymysql.connect(
    host='192.168.249.150',
    port=3306,
    user='alex',
    passwd='123456',
    db='dkh',
    charset='utf8',
    cursorclass=pymysql.cursors.SSCursor
)
cursor = conn.cursor()
cursor.execute(
    'SELECT FINALTIME, FLAG_NAME, SALENO, MEMBERCARDNO, WARENAME, WAREQTY, STDAMT, NETAMT FROM v_sale_test LIMIT 1000')
row = cursor.fetchall()
cursor.close()
conn.close()

data = pd.DataFrame(row)
data[0] = data[0].apply(lambda x: x.strftime('%Y/%m/%d %H:%M:%S'))

import sys
from termcolor import colored, cprint

text = colored('Hello, World!', 'red', attrs=['reverse', 'blink'])
print(text)
cprint('Hello, World!', 'green', 'on_red')

print_red_on_cyan = lambda x: cprint(x, 'red', 'on_cyan')
print_red_on_cyan('Hello, World!')
print_red_on_cyan('Hello, Universe!')

for i in range(10):
    cprint(i, 'magenta', end=' ')

cprint("删除【liaocheng_sale_fact库】10月份旧数据1461071行; 新增10月份数据1528956行", 'white', attrs=['bold', 'reverse', 'blink'])
print('总耗时：933秒')

from time import strftime, gmtime

print("总耗时：" + strftime("%H:%M:%S", gmtime(360)))
cprint(" > 新增门店导出成功 >> 记得发给秀清姐哦 >>> ", 'cyan', attrs=['bold', 'reverse', 'underline'])
'''
colors:
grey
red
green
yellow
blue
magenta
cyan
white

Attributes:
bold
dark
underline
blink
reverse
concealed
'''

