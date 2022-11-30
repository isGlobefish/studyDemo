# -*- coding: utf-8 -*-
'''
@createTool    : PyCharm-2020.3.3
@projectName   : pythonProjectPy3.9
@originalAuthor: Made in win10.Sys design by deHao.Zou
@createTime    : 2020/3/15 15:34
'''
import json
import math
import xlrd
import requests
import datetime
from termcolor import cprint
from time import strftime, gmtime

startTime = datetime.datetime.now()

month = input(">>>请输入上传月份：")

print(">> 读取数据中,请稍等片刻")


def excel_json_sap(filename, dataId):
    workbook = xlrd.open_workbook(filename)
    sheet = workbook.sheet_by_index(0)
    rows = sheet.nrows
    MAXROWS = 390000  # 每次只能最大上传39万行数据
    print(">> 本次总上传数据" + str(rows - 1) + "行; 需上传" + str(math.ceil(rows / MAXROWS)) + "次")
    for i in range(math.ceil(rows / MAXROWS)):
        if i == math.ceil(rows / MAXROWS) - 1:
            endRows = rows
        else:
            endRows = MAXROWS * (i + 1)
        if i == 0:
            startRows = MAXROWS * i + 1
        else:
            startRows = MAXROWS * i
        content = []
        for j in range(startRows, endRows):
            values = sheet.row_values(j)
            content.append(
                (
                    {
                        "PARTNER": str(int(values[0])),
                        "ZQDCY_FROM": values[1],
                        "MATNR_FROM": values[2],
                        "ZGGE_FROM": values[3],
                        "MENGE": values[4],
                        "MEINS_FROM": values[5],
                        "ZDATE_DEAL": str(int(values[6])),
                    }
                )

            )
        # 字典中的数据都是单引号,但是标准的json需要双引号
        Json = json.dumps(content, sort_keys=True, ensure_ascii=False, indent=4, separators=(',', ':'))
        # 前面的数据只是数组,加上外面的json格式大括号, ZNB:导入SAP的流水单号
        dataExecelJson = """{""" + """\n"ZNB": """ + """\"""" + str(dataId) + str(i + 1).zfill(2) + """\"""" + """,\n"ZDATA":""" + Json + """\n}"""
        # 用作删除指定流向单号
        # dataExecelJson = """{""" + """\n"ZNB": """ + """\"""" + str(dataId) + """\"""" + """,\n"ZDATA":""" + Json + """\n}"""
        # print(dataExecelJson)
        # 可读可写, 如果不存在则创建, 如果有内容则覆盖
        # jsFile = open("D:/FilesCenter/OnlineCustomer/jsonTest" + str(i + 1) + ".json", "w+", encoding='utf-8')
        # jsFile.write(dataExecelJson)
        # jsFile.close()

        startUpTime = datetime.datetime.now()
        print(">> 上传[" + str(MAXROWS * i + 1) + "," + str(endRows) + "]数据中,需等待时间较长")
        # r = requests.post(url, data=(json.dumps(dataJson, ensure_ascii=False)).encode("utf-8"))
        # r = requests.post(url, data=json.dumps(dataJson))
        # 正式环境(92.1/s)：http://zycrmprd01.zeus.com:8000/ZWEB/DKHLX?sap-client=800
        # 测试环境(866.86/s)：http://zycrmqas.zeus.com:8000/ZWEB/DKHLX?sap-client=600
        r = requests.post("http://zycrmprd01.zeus.com:8000/ZWEB/DKHLX?sap-client=800", data=dataExecelJson.encode("utf-8"))
        '''
        requests.get()                             # GET请求
        requests.post()                            # POST请求
        requests.put()                             # PUT请求
        requests.delete()                          # DELETE请求
        requests.head()                            # HEAD请求
        requests.options()                         # OPTIONS请求
        r.encoding                                 # 获取当前的编码
        r.encoding = 'utf-8'                       # 设置编码
        r.text                                     # 以encoding解析返回内容。字符串方式的响应体, 会自动根据响应头部的字符编码进行解码。
        r.content                                  # 以字节形式（二进制）返回。字节方式的响应体, 会自动为你解码 gzip 和 deflate 压缩。
        r.headers                                  # 以字典对象存储服务器响应头, 但是这个字典比较特殊, 字典键不区分大小写, 若键不存在则返回None
        r.status_code                              # 响应状态码
        r.raw                                      # 返回原始响应体, 也就是urllib的response对象, 使用r.raw.read()   
        r.ok                                       # 查看r.ok的布尔值便可以知道是否登陆成功
        *特殊方法*
        r.json()                                   # Requests中内置的JSON解码器,以json形式返回,前提返回的内容确保是json格式的, 不然解析出错会抛异常
        r.raise_for_status()                       # 失败请求(非200响应)抛出异常
        r.headers                                  #返回字典类型,头信息
        r.requests.headers                         #返回发送到服务器的头信息
        r.cookies                                  #返回cookie
        r.history                                  #返回重定向信息,当然可以在请求是加上allow_redirects = false 阻止重定向
        '''
        # 传出参数
        # ZMESSAGE	CHAR	200	消息文本
        # ZTYPE  	CHAR	1	消息类型: S 成功, E 错误, W 警告, I 信息, A 中断
        cprint(">> " + str(dataId) + str(i + 1).zfill(2) + ":" + str(r.json()), 'magenta', attrs=['bold', 'reverse', 'blink'])
        endUpTime = datetime.datetime.now()
        print(">> 上传[" + str(MAXROWS * i + 1) + "," + str(endRows) + "]数据耗时:" + strftime("%H:%M:%S", gmtime((endUpTime - startUpTime).seconds)))


# 注意！！！上传单号格式：DKH + 年 + 月"1100002253"
excel_json_sap("D:/FilesCenter/OnlineCustomer/DKH2021" + str(month).zfill(2) + "_OCData.xlsx", "OC2021" + str(month).zfill(2))
# excel_json_sap("D:/FilesCenter/OnlineCustomer/DKH2021" + str(month).zfill(2) + "_OCData.xlsx", "DKH20210301")

endTime = datetime.datetime.now()
print(">>>总耗时:" + strftime("%H:%M:%S", gmtime((endTime - startTime).seconds)))
