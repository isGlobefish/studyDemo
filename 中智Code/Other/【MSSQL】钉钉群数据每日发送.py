# -*- coding: utf-8 -*-
'''
@createTool    : PyCharm-2020.2.2
@projectName   : pythonProject 
@originalAuthor: Made in win10.Sys design by deHao.Zou
@createTime    : 2020/11/18 9:56
'''
import xlwt
import time
import hmac
import json
import base64
import hashlib
import pymysql
import pymssql
import requests
import calendar
import pandas as pd
import numpy as np
import urllib.parse
import urllib.request
from datetime import datetime
from time import strftime, gmtime

# 0:发送本月数据 1:发送上一个月数据
inputTpye = input(">>>请输入发送类型:")

if inputTpye == '1':  # 上个月数据
    Year = str(time.strftime("%Y", time.localtime())).zfill(4)  # 本年
    Month = str(int(time.strftime("%m", time.localtime())) - 1).zfill(2)  # 前一月
    # if today == 1:
    today = str(calendar.monthrange(int(Year), int(Month))[1]).zfill(2)  # 前一月最后一日
    # daySub2 = str(int(calendar.monthrange(int(Year), int(Month))[1]) - 1).zfill(2)  # 前一月倒数第二天
    # elif today == 2:
    # daySub1 = str(int(time.strftime("%d", time.localtime())) - 1).zfill(2)  # 前一日
    # daySub2 = str(calendar.monthrange(int(Year), int(Month))[1]).zfill(2)  # 前一月最后一日
else:  # 本月数据
    Year = str(time.strftime("%Y", time.localtime())).zfill(4)  # 本年
    Month = str(int(time.strftime("%m", time.localtime()))).zfill(2)  # 本月
    today = int(time.strftime("%d", time.localtime()))  # 本日
    # daySub1 = str(int(time.strftime("%d", time.localtime())) - 1).zfill(2)  # 前一日
    # daySub2 = str(int(time.strftime("%d", time.localtime())) - 2).zfill(2)  # 前两日


# 制作表格
def Create_TableImage(Dtable):
    # 设置单元格样式
    def set_style(fontName, height, bold=False, Halign=False, Valign=False, setBorder=False, setbgcolor=False):
        style = xlwt.XFStyle()  # 设置类型

        font = xlwt.Font()  # 为样式创建字体
        font.name = fontName
        font.height = height  # 字体大小，220就是11号字体，大概就是11*20得来
        font.bold = bold
        font.color = 'black'
        font.color_index = 4
        style.font = font

        alignment = xlwt.Alignment()  # 设置字体在单元格的位置
        if Halign == 0:
            alignment.horz = xlwt.Alignment.HORZ_CENTER  # 水平居中
        elif Halign == 1:
            alignment.horz = xlwt.Alignment.HORZ_LEFT  # 水平偏左
        else:
            alignment.horz = xlwt.Alignment.HORZ_RIGHT  # 水平偏右
        if Valign == 0:
            alignment.vert = xlwt.Alignment.VERT_CENTER  # 竖直居中
        elif Valign == 1:
            alignment.vert = xlwt.Alignment.VERT_TOP  # 竖直置顶
        else:
            alignment.vert = xlwt.Alignment.VERT_BOTTOM  # 竖直底部
        # alignment.horz = xlwt.Alignment.HORZ_CENTER  # 水平居中
        # alignment.horz = xlwt.Alignment.HORZ_LEFT  # 水平偏左
        # alignment.horz = xlwt.Alignment.HORZ_RIGHT  # 水平偏右
        # alignment.vert = xlwt.Alignment.VERT_CENTER  # 竖直居中
        # alignment.vert = xlwt.Alignment.VERT_TOP  # 竖直置顶
        # alignment.vert = xlwt.Alignment.VERT_BOTTOM  # 竖直底部
        style.alignment = alignment

        border = xlwt.Borders()  # 给单元格加框线
        if setBorder == 0:
            border.left = xlwt.Borders.THIN  # 左
            border.top = xlwt.Borders.THIN  # 上
            border.right = xlwt.Borders.THIN  # 右
            border.bottom = xlwt.Borders.THIN  # 下
            border.left_colour = 0x40  # 设置框线颜色，0x40是黑色，颜色真的巨多
            border.right_colour = 0x40
            border.top_colour = 0x40
            border.bottom_colour = 0x40
        else:
            pass
        style.borders = border

        pattern = xlwt.Pattern()  # 设置背景颜色
        if setbgcolor == 0:
            pattern.pattern = xlwt.Pattern.SOLID_PATTERN
            pattern.pattern_fore_colour = xlwt.Style.colour_map['turquoise']
        else:
            pass
        style.pattern = pattern
        return style

    try:
        # 创建表格模板并写入数据
        f = xlwt.Workbook()  # 创建工作簿
        sheet1 = f.add_sheet('Sheet1', cell_overwrite_ok=True)

        # 设置列宽
        for i in range(21):
            if i == 2 or i == 5 or i == 14:
                sheet1.col(i).width = 256 * 50
            elif i == 0:
                sheet1.col(i).width = 256 * 30
            else:
                sheet1.col(i).width = 256 * 13

        # 设置行高
        for rowi in range(0, 3):
            rowHeight = xlwt.easyxf('font:height 220;')  # 36pt,类型小初的字号
            rowNum = sheet1.row(rowi)
            rowNum.set_style(rowHeight)

        """ 设计表头 """
        colNames = ['调整区域', '商品编码', '商品名称', '商品规格', '销售日期', '门店名称', '门店编码', '数量', '单价', '零售金额', '标准单价', '标准零售金额', '客户体系',
                    '商品简称', '标识符', '大区', '省区', '实际营运区', '客户编码', '客户名称']
        # 第1行
        for icol, iname in enumerate(colNames):
            sheet1.write_merge(0, 0, icol, icol, iname,
                               set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
        # 第2行
        sheet1.write_merge(1, 1, 0, 6, "大参林",
                           set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
        sheet1.write_merge(1, 1, 7, 7,
                           str(sum([float(i) for i in list(Dtable[Dtable["客户体系"] == "大参林"].reset_index(drop=True).loc[:, '数量'])])),
                           set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
        sheet1.write_merge(1, 1, 8, 8, '',
                           set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
        sheet1.write_merge(1, 1, 9, 9,
                           str(sum([float(i) for i in list(Dtable[Dtable["客户体系"] == "大参林"].reset_index(drop=True).loc[:, '零售金额'])])),
                           set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
        sheet1.write_merge(1, 1, 10, 10, '',
                           set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
        sheet1.write_merge(1, 1, 11, 11,
                           str(sum([float(i) for i in list(Dtable[Dtable["客户体系"] == "大参林"].reset_index(drop=True).loc[:, '标准零售金额'])])),
                           set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
        # 第3行
        sheet1.write_merge(2, 2, 0, 6, "益丰",
                           set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
        sheet1.write_merge(2, 2, 7, 7,
                           str(sum([float(i) for i in list(Dtable[Dtable["客户体系"] == "益丰"].reset_index(drop=True).loc[:, '数量'])])),
                           set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
        sheet1.write_merge(2, 2, 8, 8, '',
                           set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
        sheet1.write_merge(2, 2, 9, 9,
                           str(sum([float(i) for i in list(Dtable[Dtable["客户体系"] == "益丰"].reset_index(drop=True).loc[:, '零售金额'])])),
                           set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
        sheet1.write_merge(2, 2, 10, 10, '',
                           set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
        sheet1.write_merge(2, 2, 11, 11,
                           str(sum([float(i) for i in list(Dtable[Dtable["客户体系"] == "益丰"].reset_index(drop=True).loc[:, '标准零售金额'])])),
                           set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
        # 第4行
        sheet1.write_merge(3, 3, 0, 6, "高济",
                           set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
        sheet1.write_merge(3, 3, 7, 7,
                           str(sum([float(i) for i in list(Dtable[Dtable["客户体系"] == "高济"].reset_index(drop=True).loc[:, '数量'])])),
                           set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
        sheet1.write_merge(3, 3, 8, 8, '',
                           set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
        sheet1.write_merge(3, 3, 9, 9,
                           str(sum([float(i) for i in list(Dtable[Dtable["客户体系"] == "高济"].reset_index(drop=True).loc[:, '零售金额'])])),
                           set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
        sheet1.write_merge(3, 3, 10, 10, '',
                           set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
        sheet1.write_merge(3, 3, 11, 11,
                           str(sum([float(i) for i in list(Dtable[Dtable["客户体系"] == "高济"].reset_index(drop=True).loc[:, '标准零售金额'])])),
                           set_style('等线', 210, True, Halign=0, Valign=0, setBorder=0, setbgcolor=0))
        # 填充表格数据
        for rowi in range(len(Dtable)):
            for colj, content in enumerate(colNames):
                if colj == 4:
                    sheet1.write_merge(rowi + 4, rowi + 4, colj, colj, str(Dtable.loc[rowi, content]))
                else:
                    sheet1.write_merge(rowi + 4, rowi + 4, colj, colj, Dtable.loc[rowi, content])
    except Exception as e:
        print("制作发送表格出错！！！")
    f.save('D:/FilesCenter/大客户数据/Everyday-江西/' + Year + '-' + Month + '-' + str(today).zfill(2) + '江西.xls')  # 保存文件


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
        self.webhook_url = self.roboturl + '&timestamp={}&sign={}'.format(timestamp, sign)  # 最终url,url+时间戳+签名

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


#     加一个发送文件（包括图片、文本、表格、压缩文件等等）
#     获取全部手机号（匹配人名与手机号）


if __name__ == '__main__':
    # 获取(益丰、大参林、高济)江西区域数据
    startTime = datetime.now()
    print(" > 数据获取中,请稍等片刻")
    conn = pymssql.connect(
        host='localhost',
        port=1433,
        user='sa',
        password='123456',
        database='SY',
        charset='utf8'
    )

    cursorSZ = conn.cursor()  # 海王深圳
    cursorDL = conn.cursor()  # 海王大连
    cursorJX = conn.cursor()  # 江西
    cursorHBXX = conn.cursor()  # 河北新兴
    cursorCD = conn.cursor()  # 成都
    cursorQY = conn.cursor()  # 全亿
    cursorGJ = conn.cursor()  # 高济
    cursorLBX = conn.cursor()  # 老百姓
    cursorHRT = conn.cursor()  # 惠仁堂
    cursorPZYFJH = conn.cursor()  # 普泽益丰江海
    cursorHB = conn.cursor()  # 湖北

    executeCodeSZ = """SELECT * FROM ddmonth WHERE customer = '海王' AND YEAR(date)= '""" + Year + """' AND MONTH(date)= '""" + Month + """' AND state = '深圳';"""

    executeCodeDL = """SELECT * FROM ddmonth WHERE customer = '海王' AND YEAR(date)= '""" + Year + """' AND MONTH(date)= '""" + Month + """' AND state = '大连';"""

    executeCodeJX = """SELECT * FROM ddmonth WHERE customer = '益丰' AND YEAR(date) = '""" + Year + """' AND MONTH(date) = '""" + Month + """' AND (desc_1 = '江西' ) UNION ALL
    SELECT * FROM ddmonth WHERE customer = '大参林' AND YEAR(date) = '""" + Year + """' AND MONTH(date) = '""" + Month + """' AND (desc_1 = '赣州' OR desc_1 = '南昌') UNION ALL
    SELECT * FROM ddmonth WHERE customer = '高济' AND YEAR(date) = '""" + Year + """' AND MONTH(date) = '""" + Month + """' AND (desc_1 LIKE N'%江西开心人大药房连锁有限公司%');"""

    executeCodeHBXX = """SELECT * FROM ddmonth WHERE customer = '益丰' AND YEAR(date)= '""" + Year + """' AND MONTH(date)= '""" + Month + """' AND desc_1 = '河北新兴';"""

    executeCodeCD = """SELECT * FROM ddmonth WHERE customer = '海王' AND YEAR(date)= '""" + Year + """' AND MONTH(date)= '""" + Month + """' AND desc_1 = '成都';"""

    executeCodeQY = """SELECT * FROM ddmonth WHERE customer = '全亿' AND YEAR(date)= '""" + Year + """' AND MONTH(date)= '""" + Month + """';"""

    executeCodeGJ = """SELECT * FROM ddmonth WHERE customer = '高济' AND YEAR(date)= '""" + Year + """' AND MONTH(date)= '""" + Month + """';"""

    executeCodeLBX = """SELECT * FROM ddmonth WHERE customer = '老百姓' AND YEAR(date)= '""" + Year + """' AND MONTH(date)= '""" + Month + """';"""

    executeCodeHRT = """SELECT * FROM ddmonth WHERE customer = '老百姓' AND YEAR(date)= '""" + Year + """' AND MONTH(date)= '""" + Month + """' AND desc_1 LIKE N'%惠仁堂%';"""

    executeCodePZYFJH = """SELECT * FROM ddmonth WHERE customer = '老百姓' AND YEAR(date) = '""" + Year + """' AND MONTH(date) = '""" + Month + """' AND (desc_1 LIKE N'%普泽%') UNION ALL
    SELECT * FROM ddmonth WHERE customer = '益丰' AND YEAR(date) = '""" + Year + """' AND MONTH(date) = '""" + Month + """' UNION ALL
    SELECT * FROM ddmonth WHERE customer = '大参林' AND YEAR(date) = '""" + Year + """' AND MONTH(date) = '""" + Month + """' AND (desc_1 LIKE N'%南通%');"""

    executeCodeHB = """SELECT * FROM ddmonth WHERE customer = '老百姓' AND YEAR(date) = '""" + Year + """' AND MONTH(date) = '""" + Month + """' AND (desc_1 LIKE N'%湖北%') UNION ALL
    SELECT * FROM ddmonth WHERE customer = '海王' AND YEAR(date) = '""" + Year + """' AND MONTH(date) = '""" + Month + """' AND (desc_1 LIKE N'%湖北%') UNION ALL
    SELECT * FROM ddmonth WHERE customer = '益丰' AND YEAR(date) = '""" + Year + """' AND MONTH(date) = '""" + Month + """' AND (desc_1 LIKE N'%湖北%');"""

    cursorSZ.execute(executeCodeSZ)  # 执行查询
    cursorDL.execute(executeCodeDL)  # 执行查询
    cursorJX.execute(executeCodeJX)  # 执行查询
    cursorHBXX.execute(executeCodeHBXX)  # 执行查询
    cursorCD.execute(executeCodeCD)  # 执行查询
    cursorQY.execute(executeCodeQY)  # 执行查询
    cursorGJ.execute(executeCodeGJ)  # 执行查询
    cursorLBX.execute(executeCodeLBX)  # 执行查询
    cursorHRT.execute(executeCodeHRT)  # 执行查询
    cursorPZYFJH.execute(executeCodePZYFJH)  # 执行查询
    cursorHB.execute(executeCodeHB)  # 执行查询

    rowNumSZ = cursorSZ.rowcount  # 查询数据条数
    rowNumDL = cursorDL.rowcount  # 查询数据条数
    rowNumJX = cursorJX.rowcount  # 查询数据条数
    rowNumHBXX = cursorHBXX.rowcount  # 查询数据条数
    rowNumCD = cursorCD.rowcount  # 查询数据条数
    rowNumQY = cursorQY.rowcount  # 查询数据条数
    rowNumGJ = cursorGJ.rowcount  # 查询数据条数
    rowNumLBX = cursorLBX.rowcount  # 查询数据条数
    rowNumHRT = cursorHRT.rowcount  # 查询数据条数
    rowNumPZYFJH = cursorPZYFJH.rowcount  # 查询数据条数
    rowNumHB = cursorHB.rowcount  # 查询数据条数

    dataSZ = cursorSZ.fetchall()  # 获取全部查询数据
    dataDL = cursorDL.fetchall()  # 获取全部查询数据
    dataJX = cursorJX.fetchall()  # 获取全部查询数据
    dataHBXX = cursorHBXX.fetchall()  # 获取全部查询数据
    dataCD = cursorCD.fetchall()  # 获取全部查询数据
    dataQY = cursorQY.fetchall()  # 获取全部查询数据
    dataGJ = cursorGJ.fetchall()  # 获取全部查询数据
    dataLBX = cursorLBX.fetchall()  # 获取全部查询数据
    dataHRT = cursorHRT.fetchall()  # 获取全部查询数据
    dataPZYFJH = cursorPZYFJH.fetchall()  # 获取全部查询数据
    dataHB = cursorHB.fetchall()  # 获取全部查询数据

    conn.commit()  # 提交确认

    cursorSZ.close()  # 关闭光标
    cursorDL.close()  # 关闭光标
    cursorJX.close()  # 关闭光标
    cursorHBXX.close()  # 关闭光标
    cursorCD.close()  # 关闭光标
    cursorQY.close()  # 关闭光标
    cursorGJ.close()  # 关闭光标
    cursorLBX.close()  # 关闭光标
    cursorHRT.close()  # 关闭光标
    cursorPZYFJH.close()  # 关闭光标
    cursorHB.close()  # 关闭光标

    conn.close()  # 关闭连接

    endTime = datetime.now()
    print(" > " + Year + '.' + Month + '.' + str(today).zfill(2) + "\n>> 【海王深圳】获取数据" + str(rowNumSZ) + "条;\n>> 【海王大连】获取数据" + str(
        rowNumDL) + "条;\n>> 【江西区域】获取数据" + str(rowNumJX) + "条;\n>> 【益丰河北新兴】获取数据" + str(rowNumHBXX) + "条;\n>> 【海王西南(成都)】获取数据" + str(
        rowNumCD) + "条;\n>> 【全亿】获取数据" + str(rowNumQY) + "条;\n>> 【高济】获取数据" + str(rowNumGJ) + "条;\n>> 【老百姓】获取数据" + str(
        rowNumLBX) + "条;\n>> 【惠仁堂】获取数据" + str(rowNumHRT) + "条;\n>> 【普泽益丰江海】获取数据" + str(rowNumPZYFJH) + "条;\n>> 【湖北】获取数据" + str(
        rowNumHB) + "条.\n>>>数据获取耗时：" + strftime("%H:%M:%S", gmtime((endTime - startTime).seconds)))

    print(" > 存储表格中,请稍等片刻")
    colNames = ['调整区域', '商品编码', '商品名称', '商品规格', '销售日期', '门店名称', '门店编码', '数量', '单价', '零售金额', '标准单价', '标准零售金额', '客户体系',
                '商品简称', '标识符', '大区', '省区', '实际营运区', '客户编码', '客户名称']

    shenzhenData = pd.DataFrame(dataSZ, columns=colNames, dtype=np.float64)
    shenzhenData['销售日期'] = pd.to_datetime(shenzhenData['销售日期'], format='%Y/%m/%d').dt.date
    shenzhenData.to_excel('D:/FilesCenter/大客户数据/HW-深圳/' + Year + '-' + Month + '-' + str(today).zfill(2) + '海王深圳.xlsx', index=False)

    dalianData = pd.DataFrame(dataDL, columns=colNames, dtype=np.float64)
    dalianData['销售日期'] = pd.to_datetime(dalianData['销售日期'], format='%Y/%m/%d').dt.date
    dalianData.to_excel('D:/FilesCenter/大客户数据/HW-大连/' + Year + '-' + Month + '-' + str(today).zfill(2) + '海王大连.xlsx', index=False)

    jiangxiData = pd.DataFrame(dataJX, columns=colNames, dtype=np.float64)
    jiangxiData['销售日期'] = pd.to_datetime(jiangxiData['销售日期'], format='%Y/%m/%d').dt.date
    Create_TableImage(jiangxiData)  # 创建自定义格式表格

    hebeixinxingData = pd.DataFrame(dataHBXX, columns=colNames, dtype=np.float64)
    hebeixinxingData['销售日期'] = pd.to_datetime(hebeixinxingData['销售日期'], format='%Y/%m/%d').dt.date
    hebeixinxingData.to_excel('D:/FilesCenter/大客户数据/YF-河北新兴/' + Year + '-' + Month + '-' + str(today).zfill(2) + '益丰河北新兴.xlsx', index=False)

    chengduData = pd.DataFrame(dataCD, columns=colNames, dtype=np.float64)
    chengduData['销售日期'] = pd.to_datetime(chengduData['销售日期'], format='%Y/%m/%d').dt.date
    chengduData.to_excel('D:/FilesCenter/大客户数据/HW-西南/' + Year + '-' + Month + '-' + str(today).zfill(2) + '海王西南(成都).xlsx', index=False)

    quanyiData = pd.DataFrame(dataQY, columns=colNames, dtype=np.float64)
    quanyiData['销售日期'] = pd.to_datetime(quanyiData['销售日期'], format='%Y/%m/%d').dt.date
    quanyiData.to_excel('D:/FilesCenter/大客户数据/QY-全亿/' + Year + '-' + Month + '-' + str(today).zfill(2) + '全亿.xlsx', index=False)

    gaojiData = pd.DataFrame(dataGJ, columns=colNames, dtype=np.float64)
    gaojiData['销售日期'] = pd.to_datetime(gaojiData['销售日期'], format='%Y/%m/%d').dt.date
    gaojiData.to_excel('D:/FilesCenter/大客户数据/GJ-高济/' + Year + '-' + Month + '-' + str(today).zfill(2) + '高济.xlsx', index=False)

    laobaixingData = pd.DataFrame(dataLBX, columns=colNames, dtype=np.float64)
    laobaixingData['销售日期'] = pd.to_datetime(laobaixingData['销售日期'], format='%Y/%m/%d').dt.date
    laobaixingData.to_excel('D:/FilesCenter/大客户数据/LBX-老百姓/' + Year + '-' + Month + '-' + str(today).zfill(2) + '老百姓.xlsx', index=False)

    huirentangData = pd.DataFrame(dataHRT, columns=colNames, dtype=np.float64)
    huirentangData['销售日期'] = pd.to_datetime(huirentangData['销售日期'], format='%Y/%m/%d').dt.date
    huirentangData.to_excel('D:/FilesCenter/大客户数据/LBX-惠仁堂/' + Year + '-' + Month + '-' + str(today).zfill(2) + '惠仁堂.xlsx', index=False)

    puzeyifengjianghaiData = pd.DataFrame(dataPZYFJH, columns=colNames, dtype=np.float64)
    puzeyifengjianghaiData['销售日期'] = pd.to_datetime(puzeyifengjianghaiData['销售日期'], format='%Y/%m/%d').dt.date
    puzeyifengjianghaiData.to_excel('D:/FilesCenter/大客户数据/LBX-普泽益丰江海/' + Year + '-' + Month + '-' + str(today).zfill(2) + '普泽益丰江海.xlsx', index=False)

    hubeiData = pd.DataFrame(dataHB, columns=colNames, dtype=np.float64)
    hubeiData['销售日期'] = pd.to_datetime(hubeiData['销售日期'], format='%Y/%m/%d').dt.date
    hubeiData.to_excel('D:/FilesCenter/大客户数据/Everyday-湖北/' + Year + '-' + Month + '-' + str(today).zfill(2) + '湖北.xlsx', index=False)

    # 深圳 待发送文件路径
    FilePathSZ = 'D:/FilesCenter/大客户数据/HW-深圳/' + Year + '-' + Month + '-' + str(today).zfill(2) + '海王深圳.xlsx'
    # 大连 待发送文件路径
    FilePathDL = 'D:/FilesCenter/大客户数据/HW-大连/' + Year + '-' + Month + '-' + str(today).zfill(2) + '海王大连.xlsx'
    # 江西 待发送文件路径
    FilePathJX = 'D:/FilesCenter/大客户数据/Everyday-江西/' + Year + '-' + Month + '-' + str(today).zfill(2) + '江西.xls'
    # 河北新兴 待发送文件路径
    FilePathHBXX = 'D:/FilesCenter/大客户数据/YF-河北新兴/' + Year + '-' + Month + '-' + str(today).zfill(2) + '益丰河北新兴.xlsx'
    # 成都 待发送文件路径
    FilePathCD = 'D:/FilesCenter/大客户数据/HW-西南/' + Year + '-' + Month + '-' + str(today).zfill(2) + '海王西南(成都).xlsx'
    # 全亿 待发送文件路径
    FilePathQY = 'D:/FilesCenter/大客户数据/QY-全亿/' + Year + '-' + Month + '-' + str(today).zfill(2) + '全亿.xlsx'
    # 高济 待发送文件路径
    FilePathGJ = 'D:/FilesCenter/大客户数据/GJ-高济/' + Year + '-' + Month + '-' + str(today).zfill(2) + '高济.xlsx'
    # 老百姓 待发送文件路径
    FilePathLBX = 'D:/FilesCenter/大客户数据/LBX-老百姓/' + Year + '-' + Month + '-' + str(today).zfill(2) + '老百姓.xlsx'
    # 惠仁堂 待发送文件路径
    FilePathHRT = 'D:/FilesCenter/大客户数据/LBX-惠仁堂/' + Year + '-' + Month + '-' + str(today).zfill(2) + '惠仁堂.xlsx'
    # 普泽益丰江海 待发送文件路径
    FilePathPZYFJH = 'D:/FilesCenter/大客户数据/LBX-普泽益丰江海/' + Year + '-' + Month + '-' + str(today).zfill(2) + '普泽益丰江海.xlsx'
    # 湖北 待发送文件路径
    FilePathHB = 'D:/FilesCenter/大客户数据/Everyday-湖北/' + Year + '-' + Month + '-' + str(today).zfill(2) + '湖北.xlsx'

    # 通过jsapi工具 https://wsdebug.dingtalk.com/?spm=a219a.7629140.0.0.7bc84a972WUfGd 获取目标群聊chatId
    ChatIdSZ = 'chat36cbc03177f4b48623d505f449852499'  # 深圳
    ChatIdDL = 'chatca0d82c9d6e94531f7d639b33653a8f5'  # 大连
    ChatIdJX = 'chat8971c7614a674dc958a6a4baf91f2632'  # 江西
    ChatIdHBXX = 'chat8e8f0ba98e0e512377716d416bf992a5'  # 河北新兴
    ChatIdXN = 'chat75495b5797bc775a06bd9722f42a8260'  # 西南
    ChatIdLBX = 'chat34ffac1b8839a6a3df8854750c9a8836'  # 老百姓
    ChatIdHRT = 'chat0cb3560baecbad4f68938a3d872f5a0c'  # 惠仁堂
    ChatIdPZYFJH = 'chatcf38482868ce95eed23540f7e3a6820d'  # 普泽益丰江海
    ChatIdHB = 'chat9bee77b820c60cbee4cc49877ec19596'  # 湖北

    AppKey = 'dingjpjkc2vaqjoqgmhz'  # 企业开发平台小程序AppKey
    AppSecret = 'oKNcuSF12oW0j9eBeO53wA6qwmKCVz34NVy1NvtvnjsvKPOdKiozsSZzUypNSWDc'  # 企业开发平台小程序AppSecret

    RobotWebHookURLSZ = 'https://oapi.dingtalk.com/robot/send?access_token=bbffc889ad9051e4ea7d011f46c15146e7ab893f1cccd2b2a9170d7c54e47de0'  # 深圳群机器人url
    RobotWebHookURLDL = 'https://oapi.dingtalk.com/robot/send?access_token=262aae91ab9db9d928bc8c933f1136b139582596a34a5e315e7916d362abc5fc'  # 大连群机器人url
    RobotWebHookURLJX = 'https://oapi.dingtalk.com/robot/send?access_token=ad57e60ff8079e27cde8b59d06c0ef46303e6235eedfcab03ab368757d6a468a'  # 江西群机器人url
    RobotWebHookURLHBXX = 'https://oapi.dingtalk.com/robot/send?access_token=d3919b5b65783f24e5bf78b2ce1fb79e7eefa043ca1944eff7b683d8e6fed22f'  # 河北新兴群机器人url
    RobotWebHookURLXN = 'https://oapi.dingtalk.com/robot/send?access_token=89cbf9cfca724a54cf8939cf5e10fdb013a4581554716584fa189c00f2779583'  # 西南群机器人url
    RobotWebHookURLLBX = 'https://oapi.dingtalk.com/robot/send?access_token=6dabd3db1800f83198f07a228d619e62a46f70dd5d58e7ddc7770f036ea81ce8'  # 老百姓群机器人url
    RobotWebHookURLHRT = 'https://oapi.dingtalk.com/robot/send?access_token=88861e201507a056ba9582f15fee29d41826fa0acb21160c20044f11f2ede71a'  # 惠仁堂群机器人url
    RobotWebHookURLPZYFJH = 'https://oapi.dingtalk.com/robot/send?access_token=ec14984eae07bd0d68a8664151a053718569eb7671abfbfe7c8ce2fc0df98aa2'  # 普泽益丰江海群机器人url
    RobotWebHookURLHB = 'https://oapi.dingtalk.com/robot/send?access_token=7e0ecf1bafb8894a6181e8e6ece961dec1bf5f57e9a7fe54bb528bdc3931c233'  # 湖北群机器人url

    RobotSecret = 'GbSFeeIHgYNJfXT5WoPT6c6GRmMVRd2wVODyexo7SQIF5HJkucowab6cNMiyR8IV'  # 群机器人加签秘钥secret(默认 数运小助手 )

    ddMessageSZ = {  # 深圳消息内容
        "msgtype": "markdown",
        "markdown": {"title": "海王深圳每日数据",  # @某人 才会显示标题
                     "text": "> ###### **大家好！我是 数运小助手 机器人。今日数据已经送达, 请大家注意查收哦(⊙o⊙), 如有疑问请及时回馈, 感谢您们的理解与支持。**"
                             "\n> ###### **特别说明：海王、大参林数据延后2天, 其他客户延后1天！本数据为网上客户流向数据, 与实时销售数据可能存在差异, 考核、业绩、提成等不建议使用该数据, 建议用纯销回流向数据。**"
                             "\n###### *温馨提示：<文件>需下载才能浏览, 不能在线浏览哦！！！*"
                             "\n###### ----------------------------------------------"
                             "\n###### 数据来源：中智药业集团之数据运营中心"
                             "\n###### 发布时间：" + str(datetime.now()).split('.')[0]},  # 发布时间
        "at": {
            # "atMobiles": [15817552982],  # 指定@某人
            "isAtAll": False  # 是否@所有人[False:否, True:是]
        }
    }

    ddMessageDL = {  # 大连消息内容
        "msgtype": "markdown",
        "markdown": {"title": "海王大连每日数据",  # @某人 才会显示标题
                     "text": "> ###### **大家好！我是 数运小助手 机器人。今日数据已经送达, 请大家注意查收哦(⊙o⊙), 如有疑问请及时回馈, 感谢您们的理解与支持。**"
                             "\n> ###### **特别说明：海王、大参林数据延后2天, 其他客户延后1天！本数据为网上客户流向数据, 与实时销售数据可能存在差异, 考核、业绩、提成等不建议使用该数据, 建议用纯销回流向数据。**"
                             "\n###### *温馨提示：<文件>需下载才能浏览, 不能在线浏览哦！！！*"
                             "\n###### ----------------------------------------------"
                             "\n###### 数据来源：中智药业集团之数据运营中心"
                             "\n###### 发布时间：" + str(datetime.now()).split('.')[0]},  # 发布时间
        "at": {
            # "atMobiles": [15817552982],  # 指定@某人
            "isAtAll": False  # 是否@所有人[False:否, True:是]
        }
    }

    ddMessageJX = {  # 江西消息内容
        "msgtype": "markdown",
        "markdown": {"title": "江西区域每日数据",  # @某人 才会显示标题
                     "text": "> ###### **大家好！我是 数运小助手 机器人。今日数据已经送达, 请大家注意查收哦(⊙o⊙), 如有疑问请及时回馈, 感谢您们的理解与支持。**"
                             "\n> ###### **特别说明：海王、大参林数据延后2天, 其他客户延后1天！本数据为网上客户流向数据, 与实时销售数据可能存在差异, 考核、业绩、提成等不建议使用该数据, 建议用纯销回流向数据。**"
                             "\n###### *温馨提示：<文件>需下载才能浏览, 不能在线浏览哦！！！*"
                             "\n###### ----------------------------------------------"
                             "\n###### 数据来源：中智药业集团之数据运营中心"
                             "\n###### 发布时间：" + str(datetime.now()).split('.')[0]},  # 发布时间
        "at": {
            # "atMobiles": [15817552982],  # 指定@某人
            "isAtAll": False  # 是否@所有人[False:否, True:是]
        }
    }

    ddMessageHBXX = {  # 益丰河北新兴消息内容
        "msgtype": "markdown",
        "markdown": {"title": "益丰河北新兴每日数据",  # @某人 才会显示标题
                     "text": "> ###### **大家好！我是 数运小助手 机器人。今日数据已经送达, 请大家注意查收哦(⊙o⊙), 如有疑问请及时回馈, 感谢您们的理解与支持。**"
                             "\n> ###### **特别说明：海王、大参林数据延后2天, 其他客户延后1天！本数据为网上客户流向数据, 与实时销售数据可能存在差异, 考核、业绩、提成等不建议使用该数据, 建议用纯销回流向数据。**"
                             "\n###### *温馨提示：<文件>需下载才能浏览, 不能在线浏览哦！！！*"
                             "\n###### ----------------------------------------------"
                             "\n###### 数据来源：中智药业集团之数据运营中心"
                             "\n###### 发布时间：" + str(datetime.now()).split('.')[0]},  # 发布时间
        "at": {
            # "atMobiles": [15817552982],  # 指定@某人
            "isAtAll": False  # 是否@所有人[False:否, True:是]
        }
    }

    ddMessageXN = {  # 海王西南消息内容
        "msgtype": "markdown",
        "markdown": {"title": "海王西南区域每日数据",  # @某人 才会显示标题
                     "text": "> ###### **大家好！我是 数运小助手 机器人。今日数据已经送达, 请大家注意查收哦(⊙o⊙), 如有疑问请及时回馈, 感谢您们的理解与支持。**"
                             "\n> ###### **特别说明：海王、大参林数据延后2天, 其他客户延后1天！本数据为网上客户流向数据, 与实时销售数据可能存在差异, 考核、业绩、提成等不建议使用该数据, 建议用纯销回流向数据。**"
                             "\n###### *温馨提示：<文件>需下载才能浏览, 不能在线浏览哦！！！*"
                             "\n###### ----------------------------------------------"
                             "\n###### 数据来源：中智药业集团之数据运营中心"
                             "\n###### 发布时间：" + str(datetime.now()).split('.')[0]},  # 发布时间
        "at": {
            # "atMobiles": [15817552982],  # 指定@某人
            "isAtAll": False  # 是否@所有人[False:否, True:是]
        }
    }

    ddMessageLBX = {  # 老百姓消息内容
        "msgtype": "markdown",
        "markdown": {"title": "老百姓每日数据",  # @某人 才会显示标题
                     "text": "> ###### **大家好！我是 数运小助手 机器人。今日数据已经送达, 请大家注意查收哦(⊙o⊙), 如有疑问请及时回馈, 感谢您们的理解与支持。**"
                             "\n> ###### **特别说明：海王、大参林数据延后2天, 其他客户延后1天！本数据为网上客户流向数据, 与实时销售数据可能存在差异, 考核、业绩、提成等不建议使用该数据, 建议用纯销回流向数据。**"
                             "\n###### *温馨提示：<文件>需下载才能浏览, 不能在线浏览哦！！！*"
                             "\n###### ----------------------------------------------"
                             "\n###### 数据来源：中智药业集团之数据运营中心"
                             "\n###### 发布时间：" + str(datetime.now()).split('.')[0]},  # 发布时间
        "at": {
            # "atMobiles": [15817552982],  # 指定@某人
            "isAtAll": False  # 是否@所有人[False:否, True:是]
        }
    }

    ddMessageHRT = {  # 惠仁堂消息内容
        "msgtype": "markdown",
        "markdown": {"title": "惠仁堂每日数据",  # @某人 才会显示标题
                     "text": "> ###### **大家好！我是 数运小助手 机器人。今日数据已经送达, 请大家注意查收哦(⊙o⊙), 如有疑问请及时回馈, 感谢您们的理解与支持。**"
                             "\n> ###### **特别说明：海王、大参林数据延后2天, 其他客户延后1天！本数据为网上客户流向数据, 与实时销售数据可能存在差异, 考核、业绩、提成等不建议使用该数据, 建议用纯销回流向数据。**"
                             "\n###### *温馨提示：<文件>需下载才能浏览, 不能在线浏览哦！！！*"
                             "\n###### ----------------------------------------------"
                             "\n###### 数据来源：中智药业集团之数据运营中心"
                             "\n###### 发布时间：" + str(datetime.now()).split('.')[0]},  # 发布时间
        "at": {
            # "atMobiles": [15817552982],  # 指定@某人
            "isAtAll": False  # 是否@所有人[False:否, True:是]
        }
    }

    ddMessagePZYFJH = {  # 普泽益丰江海消息内容
        "msgtype": "markdown",
        "markdown": {"title": "普泽每日数据",  # @某人 才会显示标题
                     "text": "> ###### **大家好！我是 数运小助手 机器人。今日数据已经送达, 请大家注意查收哦(⊙o⊙), 如有疑问请及时回馈, 感谢您们的理解与支持。**"
                             "\n> ###### **特别说明：海王、大参林数据延后2天, 其他客户延后1天！本数据为网上客户流向数据, 与实时销售数据可能存在差异, 考核、业绩、提成等不建议使用该数据, 建议用纯销回流向数据。**"
                             "\n###### *温馨提示：<文件>需下载才能浏览, 不能在线浏览哦！！！*"
                             "\n###### ----------------------------------------------"
                             "\n###### 数据来源：中智药业集团之数据运营中心"
                             "\n###### 发布时间：" + str(datetime.now()).split('.')[0]},  # 发布时间
        "at": {
            # "atMobiles": [15817552982],  # 指定@某人
            "isAtAll": False  # 是否@所有人[False:否, True:是]
        }
    }

    ddMessageHB = {  # 湖北消息内容
        "msgtype": "markdown",
        "markdown": {"title": "湖北每日数据",  # @某人 才会显示标题
                     "text": "> ###### **大家好！我是 数运小助手 机器人。今日数据已经送达, 请大家注意查收哦(⊙o⊙), 如有疑问请及时回馈, 感谢您们的理解与支持。**"
                             "\n> ###### **特别说明：海王、大参林数据延后2天, 其他客户延后1天！本数据为网上客户流向数据, 与实时销售数据可能存在差异, 考核、业绩、提成等不建议使用该数据, 建议用纯销回流向数据。**"
                             "\n###### *温馨提示：<文件>需下载才能浏览, 不能在线浏览哦！！！*"
                             "\n###### ----------------------------------------------"
                             "\n###### 数据来源：中智药业集团之数据运营中心"
                             "\n###### 发布时间：" + str(datetime.now()).split('.')[0]},  # 发布时间
        "at": {
            # "atMobiles": [15817552982],  # 指定@某人
            "isAtAll": False  # 是否@所有人[False:否, True:是]
        }
    }

    """
    特别说明：
            发送消息：目前支持text、link、markdown等形式文字及图片，并不支持本地文件和图片类媒体文件的发送
            发送文件：目前支持简单excel表(csv、xlsx、xls等)、word、压缩文件,不支持ppt等文件的发送
    """

    # 深圳海王数据群
    dingdingFunction(RobotWebHookURLSZ, RobotSecret, AppKey, AppSecret).sendFile(ChatIdSZ, FilePathSZ)  # 发送文件
    dingdingFunction(RobotWebHookURLSZ, RobotSecret, AppKey, AppSecret).sendMessage(ddMessageSZ)  # 发送消息

    # 大连海王流向沟通群
    dingdingFunction(RobotWebHookURLDL, RobotSecret, AppKey, AppSecret).sendFile(ChatIdDL, FilePathDL)  # 发送文件
    dingdingFunction(RobotWebHookURLDL, RobotSecret, AppKey, AppSecret).sendMessage(ddMessageDL)  # 发送消息

    # 草晶华江西区域数据发送
    dingdingFunction(RobotWebHookURLJX, RobotSecret, AppKey, AppSecret).sendFile(ChatIdJX, FilePathJX)  # 发送文件
    dingdingFunction(RobotWebHookURLJX, RobotSecret, AppKey, AppSecret).sendMessage(ddMessageJX)  # 发送消息

    # 河北新兴流向沟通群
    dingdingFunction(RobotWebHookURLHBXX, RobotSecret, AppKey, AppSecret).sendFile(ChatIdHBXX, FilePathHBXX)  # 发送文件
    dingdingFunction(RobotWebHookURLHBXX, RobotSecret, AppKey, AppSecret).sendMessage(ddMessageHBXX)  # 发送消息

    # 海王西南大区数据群
    dingdingFunction(RobotWebHookURLXN, RobotSecret, AppKey, AppSecret).sendFile(ChatIdXN, FilePathCD)  # 发送文件
    dingdingFunction(RobotWebHookURLXN, RobotSecret, AppKey, AppSecret).sendFile(ChatIdXN, FilePathQY)  # 发送文件
    dingdingFunction(RobotWebHookURLXN, RobotSecret, AppKey, AppSecret).sendFile(ChatIdXN, FilePathGJ)  # 发送文件
    dingdingFunction(RobotWebHookURLXN, RobotSecret, AppKey, AppSecret).sendMessage(ddMessageXN)  # 发送消息

    # 老百姓数据流向群
    dingdingFunction(RobotWebHookURLLBX, RobotSecret, AppKey, AppSecret).sendFile(ChatIdLBX, FilePathLBX)  # 发送文件
    dingdingFunction(RobotWebHookURLLBX, RobotSecret, AppKey, AppSecret).sendMessage(ddMessageLBX)  # 发送消息

    # 惠仁堂数据流向群
    dingdingFunction(RobotWebHookURLHRT, RobotSecret, AppKey, AppSecret).sendFile(ChatIdHRT, FilePathHRT)  # 发送文件
    dingdingFunction(RobotWebHookURLHRT, RobotSecret, AppKey, AppSecret).sendMessage(ddMessageHRT)  # 发送消息

    # 普泽益丰江海数据群
    dingdingFunction(RobotWebHookURLPZYFJH, RobotSecret, AppKey, AppSecret).sendFile(ChatIdPZYFJH, FilePathPZYFJH)  # 发送文件
    dingdingFunction(RobotWebHookURLPZYFJH, RobotSecret, AppKey, AppSecret).sendMessage(ddMessagePZYFJH)  # 发送消息

    # 湖北老百姓益丰海王数据群
    dingdingFunction(RobotWebHookURLHB, RobotSecret, AppKey, AppSecret).sendFile(ChatIdHB, FilePathHB)  # 发送文件
    dingdingFunction(RobotWebHookURLHB, RobotSecret, AppKey, AppSecret).sendMessage(ddMessageHB)  # 发送消息
