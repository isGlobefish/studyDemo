# -*- coding: utf-8 -*-
'''
@createTool    : PyCharm-2020.2.2
@projectName   : pythonProject 
@originalAuthor: Made in win10.Sys design by deHao.Zou
@createTime    : 2020/11/11 15:20
'''
import datetime
import json
import urllib.request
import pymysql


def getMessageType():
    # 判断当天是周几选择出文案的函数
    # 获取当天日期
    today = datetime.date.today()
    # 获取当天是周几
    todayweek = datetime.date.isoweekday(today)
    # 利用IF语句判断周几选出当天要发送的文案
    if todayweek == 1:
        Copywriting = "### 每日数据  \n > 所有的成绩都始于默默搬砖！昨天的成交达到**%s**美金啦~感谢所有默默搬砖的你们，今天还是要以正能量的姿态迎接更大的挑战，加油，小伙伴们\n\n > ![screenshot](https://images-global.kikuu.com/upload-productImg-1535016385314.jpeg)\n  > ##### 10点00分发布 [BI部门](http://data.kikuu.com:8007/dashboard/?project=default#dashid=28) "
    elif todayweek == 2:
        Copywriting = "### 每日数据  \n > 如果有些事无法回避，那我们能做的，就是把自己变得更强大，强大到能够应对这一次挑战。送走昨日**%s**美金成交的历史，今日又是富有挑战的一天~\n\n > ![screenshot](https://images-global.kikuu.com/upload-productImg-1535016385314.jpeg)\n  > ##### 10点00分发布 [BI部门](http://data.kikuu.com:8007/dashboard/?project=default#dashid=28) "
    elif todayweek == 3:
        Copywriting = "### 每日数据  \n > 努力是人生的一种精神状态，往往最美的不是成功的那一刻，而是那段努力奋斗的过程。伙伴们，昨日又是一个漂亮的翻身仗，成交**%s**美金啦，愿你努力后的今天更精彩。早安！\n\n > ![screenshot](https://unsplash.com/photos/rYWKAgO7jQg)\n  > ##### 10点00分发布 [BI部门](http://data.kikuu.com:8007/dashboard/?project=default#dashid=28) "
    elif todayweek == 4:
        Copywriting = "### 每日数据  \n > 昨日的辛勤劳作又有了新突破，昨日成交已经**%s**美金啦~不抛弃不放弃，没有办法的时候，死磕也是种办法。\n\n > ![screenshot](https://images-global.kikuu.com/upload-productImg-1535019037528.jpeg)\n  > ##### 10点00分发布 [BI部门](http://data.kikuu.com:8007/dashboard/?project=default#dashid=28) "
    elif todayweek == 5:
        Copywriting = "### 每日数据  \n > 明天就是周末了，嘘~~~不要笑出声。昨天平台成交**%s**美金恩，现在可以笑出来了。又是新的一天，加油。\n\n > ![screenshot](https://images-global.kikuu.com/upload-productImg-1535016385314.jpeg)\n  > ##### 10点00分发布 [BI部门](http://data.kikuu.com:8007/dashboard/?project=default#dashid=28) "
    elif todayweek == 6:
        Copywriting = "### 每日数据  \n > 辛勤的付出才能得到我们想要的回报，一味的幻想，只会让你离梦想越来越远。看，我们的梦想又近了一步，昨日已经**%s**美金啦，成功已越来越近啦~\n\n > ![screenshot](https://images-global.kikuu.com/upload-productImg-1535016385314.jpeg)\n  > ##### 10点00分发布 [BI部门](http://data.kikuu.com:8007/dashboard/?project=default#dashid=28) "
    elif todayweek == 7:
        Copywriting = "### 每日数据  \n > 把弯路走直的人是聪明的，因为找到了捷径；把直路走弯的人是豁达的，因为可以多看几道风景；路不在脚下，路在心里。告诉大家一个好消息，昨日成交**%s**美金啦，各位早安，愿好。\n\n > ![screenshot](https://images-global.kikuu.com/upload-productImg-1535016385314.jpeg)\n  > ##### 10点00分发布 [BI部门](http://data.kikuu.com:8007/dashboard/?project=default#dashid=28) "
    return Copywriting


def send_request(url, datas):
    # 传入url和内容发送请求
    # 构建一下请求头部
    header = {
        "Content-Type": "application/json",
        "Charset": "UTF-8"
    }
    sendData = json.dumps(datas)  # 将字典类型数据转化为json格式
    sendDatas = sendData.encode("utf-8")  # python3的Request要求data为byte类型
    # 发送请求
    request = urllib.request.Request(url=url, data=sendDatas, headers=header)
    # 将请求发回的数据构建成为文件格式
    opener = urllib.request.urlopen(request)
    # 7、打印返回的结果
    print(opener.read())


def connSQL(sql):
    # 一个传入sql导出数据的函数
    # 跟数据库建立连接
    conn = pymysql.connect(host='实例地址', user='用户名',
                           passwd='密码', database='库名', port=3306, charset="utf8")
    # 使用 cursor() 方法创建一个游标对象 cursor
    cur = conn.cursor()
    # 使用 execute() 方法执行 SQL
    cur.execute(sql)
    # 获取所需要的数据
    datas = cur.fetchall()
    # 关闭连接
    cur.close()
    # 返回所需的数据
    return datas


def ddSendMessage():
    # 按照钉钉给的数据格式设计请求内容  链接https://open-doc.dingtalk.com/docs/doc.htm?spm=a219a.7629140.0.0.p7hJKp&treeId=257&articleId=105735&docType=1
    message = {
        "msgtype": "markdown",
        "markdown": {"title": "每日早报",
                     "text": " "
                     },
        "at": {

            "isAtAll": True
        }
    }
    # 获取当天文案
    sendType = getMessageType()
    # 获取昨日成交
    my_mydata = connSQL(
        "SELECT sum(usdAmount) FROM dplus_source_productorder_v2 WHERE RealPaidTime >= '2018-08-20 00:00:00' AND RealPaidTime <= '2018-08-20 23:59:59'")
    # 获取昨日成交的数值
    my_mydata = my_mydata[0][0]
    # 保留2位小数
    my_mydata = "%.2f" % my_mydata
    # 把文案中的金额替换为昨天成交金额
    my_Copywriting = sendType % my_mydata
    # 把文案内容写入请求格式中
    message["markdown"]["text"] = my_Copywriting
    # 你的钉钉机器人url
    ddURL = "复制钉钉你的机器人url地址"
    send_request(ddURL, message)


if __name__ == "__main__":
    ddSendMessage()
