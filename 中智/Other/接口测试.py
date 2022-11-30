# -*- coding: utf-8 -*-
'''
@createTool    : PyCharm-2020.3.3
@projectName   : pythonProjectPy3.9
@originalAuthor: Made in win10.Sys design by deHao.Zou
@createTime    : 2020/3/15 15:34
'''
'''
# -Begin-----------------------------------------------------------------

# -Packages--------------------------------------------------------------
from ctypes import *


# -Structures------------------------------------------------------------
class RFC_ERROR_INFO(Structure):
    _fields_ = [("code", c_long),
                ("group", c_long),
                ("key", c_wchar * 128),
                ("message", c_wchar * 512),
                ("abapMsgClass", c_wchar * 21),
                ("abapMsgType", c_wchar * 2),
                ("abapMsgNumber", c_wchar * 4),
                ("abapMsgV1", c_wchar * 51),
                ("abapMsgV2", c_wchar * 51),
                ("abapMsgV3", c_wchar * 51),
                ("abapMsgV4", c_wchar * 51)]


class RFC_CONNECTION_PARAMETER(Structure):
    _fields_ = [("name", c_wchar_p),
                ("value", c_wchar_p)]


# -Main------------------------------------------------------------------
ErrInf = RFC_ERROR_INFO;
RfcErrInf = ErrInf()
ConnParams = RFC_CONNECTION_PARAMETER * 5;
RfcConnParams = ConnParams()

SAPNWRFC = "sapnwrfc.dll"
SAP = windll.LoadLibrary(SAPNWRFC)

# -Prototypes------------------------------------------------------------
SAP.RfcOpenConnection.argtypes = [POINTER(ConnParams), c_ulong, POINTER(ErrInf)]
SAP.RfcOpenConnection.restype = c_void_p
SAP.RfcCloseConnection.argtypes = [c_void_p, POINTER(ErrInf)]
SAP.RfcCloseConnection.restype = c_ulong

# -Connection parameters-------------------------------------------------
RfcConnParams[0].name = "ASHOST";
RfcConnParams[0].value = "ABAP"
RfcConnParams[1].name = "SYSNR";
RfcConnParams[1].value = "00"
RfcConnParams[2].name = "CLIENT";
RfcConnParams[2].value = "001"
RfcConnParams[3].name = "USER";
RfcConnParams[3].value = "BCUSER"
RfcConnParams[4].name = "PASSWD";
RfcConnParams[4].value = "minisap"

hRFC = SAP.RfcOpenConnection(RfcConnParams, 5, RfcErrInf)
if hRFC != None:

    windll.user32.MessageBoxW(None, "Check connection with TAC SMGW", "", 0)

    # ---------------------------------------------------------------------
    # -
    # - Check connection with TAC SMGW in the SAP system
    # -
    # ---------------------------------------------------------------------

    rc = SAP.RfcCloseConnection(hRFC, RfcErrInf)

else:
    print(RfcErrInf.key)
    print(RfcErrInf.message)

# -End-------------------------------------------------------------------
'''
# -*- coding: utf-8 -*-
from flask import Flask, request
import json

app = Flask(__name__)


@app.route('/hci', methods=["POST"])
def check():
    data = request.get_data()
    url = json.loads(data)
    result_ = []
    score_datas = []
    temp_1 = {}
    for i in range(len(url)):
        answer = url["data{}".format(i + 1)]
        id = answer["id"]
        q1 = answer["q1"]
        q2 = answer["q2"]
        score = Similarity(q1, q2)

        temp_2 = {}
        temp_2["id"] = str(id)
        temp_2["score"] = str(score)
        result_.append(temp_2)
    total_score = get_Total_score()
    temp_1["TotalScore"] = str(total_score)
    temp_1["data"] = result_
    score_datas.append(temp_1)
    score_datas = json.dumps(score_datas, ensure_ascii=False)
    return score_datas


def Similarity(str1, str2):
    _score = 80
    return _score


def get_Total_score():
    _score = 80
    return _score


if __name__ == '__main__':
    app.run()

# ======================================================================================

# 把Excel数据转换为Json文件
import xlrd, json


def read_xlsx_file(filename):
    # 打开Excel文件
    data = xlrd.open_workbook(filename)
    # 读取第一个工作表
    table = data.sheets()[0]
    # 统计行数
    rows = table.nrows
    data = []  # 存放数据
    for i in range(1, rows):
        values = table.row_values(i)
        data.append(
            (
                {
                    "业务伙伴编码": values[0],
                    "流向门店名称": values[1],
                    "原始物料描述": values[2],
                    "原始物料规格": values[3],
                    "数量": values[4],
                    "单位": values[5],
                    "交易日期": values[6],
                }
            )

        )
    return data


if __name__ == '__main__':
    d1 = read_xlsx_file("C:/Users/Zeus/Desktop/网上客户流向3月份数据.xlsx")
    # 字典中的数据都是单引号，但是标准的json需要双引号
    js = json.dumps(d1, sort_keys=True, ensure_ascii=False, indent=4, separators=(',', ':'))
    print(js)
    # 前面的数据只是数组，加上外面的json格式大括号
    js = "{" + js + "}"
    # 可读可写，如果不存在则创建，如果有内容则覆盖
    jsFile = open("C:/Users/Zeus/Desktop/text3.json", "w+", encoding='utf-8')
    jsFile.write(js)
    jsFile.close()

# =========================================================================================
# 模拟postman上传数据
import requests

files = {'skFile': open("C:/Users/Zeus/Desktop/test.json", 'rb')}

r = requests.post("http://192.168.20.207:8000/wm/obtain?sap-client=600", files=files)

from urllib3 import encode_multipart_formdata
import requests


def post_files(url, header, data, filename, filepath):
    """
        :param files: (optional) Dictionary of ``'name': file-like-objects`` (or ``{'name': file-tuple}``) for multipart encoding upload.
        ``file-tuple`` can be a 2-tuple ``('filename', fileobj)``, 3-tuple ``('filename', fileobj, 'content_type')``
        or a 4-tuple ``('filename', fileobj, 'content_type', custom_headers)``, where ``'content-type'`` is a string
        defining the content type of the given file and ``custom_headers`` a dict-like object containing additional headers
        to add for the file.
    """
    data['file'] = (filename, open(filepath, 'rb').read())
    encode_data = encode_multipart_formdata(data)
    data = encode_data[0]
    header['Content-Type'] = encode_data[1]
    r = requests.post(url, headers=header, data=data)
    print(r.content)


if __name__ == "__main__":
    # url,filename,filepath string
    # header,data dict
    print(post_files("url", {"header": "value"}, {"data": "value"}, "filename", "filepath"))

# ======================================================================================

from urllib3 import encode_multipart_formdata
import requests

url = "http://192.168.20.207:8000/wm/obtain?sap-client=600"
data = {}
headers = {}
filename = 'JsonFiles'  # 上传至服务器后，用于存储文件的名称
filepath = 'C:/Users/Zeus/Desktop/text3.json'  # 当前要上传的文件路径
proxies = {
    "http": "http://192.168.20.207:8000",
    "https": "http://192.168.20.207:8000",
    # 如果代理ip需要用户名和密码的话 'http':'user:password@192.168.1.1:88'
}
####
data['upload_file'] = (filename, open(filepath, 'rb').read())
data['submit'] = "提交"
encode_data = encode_multipart_formdata(data)
data = encode_data[0]
headers['Content-Type'] = encode_data[1]
# r = requests.post(url, headers=headers, data=data, timeout=5)
r = requests.post(url, headers=headers, data=data, proxies=proxies, timeout=5)
print(r.status_code)

# ======================================================================================

import requests
import json
import xlrd

host = "http://httpbin.org/"
endpoint = "post"
url = ''.join([host, endpoint])
url = "http://zycrmqas.zeus.com:8000/ZWEB/DKHLX?sap-client=600"

dataJson = {
    "sites": [
        {"name": "test", "url": "www.test.com"},
        {"name": "google", "url": "www.google.com"},
        {"name": "weibo", "url": "www.weibo.com"}
    ]
}
dataJson = {
    "ZNB": "123123123",
    "ZDATA": [
        {
            "PARTNER": "111111111",
            "ZQDCY_FROM": "老百姓",
            "MATNR_FROM": "302020202",
            "ZGGE_FROM": "32g",
            "MENGE": "32",
            "MEINS_FROM": "罐",
            "ZDATE_DEAL": "20210101"

        }, {
            "PARTNER": "111111111",
            "ZQDCY_FROM": "老百姓",
            "MATNR_FROM": "302020202",
            "ZGGE_FROM": "32g",
            "MENGE": "32",
            "MEINS_FROM": "罐",
            "ZDATE_DEAL": "20210101"

        }
    ]
}


def read_xlsx_file(filename):
    # 打开Excel文件
    data = xlrd.open_workbook(filename)
    # 读取第一个工作表
    table = data.sheets()[0]
    # 统计行数
    rows = table.nrows
    data = []  # 存放数据
    for i in range(1, rows):
        values = table.row_values(i)
        data.append(
            (
                {
                    "业务伙伴编码": str(int(values[0])),
                    "流向门店名称": values[1],
                    "原始物料描述": values[2],
                    "原始物料规格": values[3],
                    "数量": values[4],
                    "单位": values[5],
                    "交易日期": str(int(values[6])),
                }
            )

        )
    return data


dataExecel = read_xlsx_file("C:/Users/Zeus/Desktop/网上客户流向3月份数据.xlsx")
# 字典中的数据都是单引号,但是标准的json需要双引号
dataJson = json.dumps(dataExecel, sort_keys=True, ensure_ascii=False, indent=4, separators=(',', ':'))
# 前面的数据只是数组,加上外面的json格式大括号
dataJson = "{" + dataJson + "}"
print(dataJson)

'''
requests.get()                                # GET请求
requests.post()                               # POST请求
requests.put()                                # PUT请求
requests.delete()                             # DELETE请求
requests.head()                               # HEAD请求
requests.options()                            # OPTIONS请求
'''
# r = requests.get('https://github.com/Ranxf')  # 最基本的不带参数的get请求
# r = requests.get(url='http://dict.baidu.com/s', params={'wd': 'python'})  # 带参数的get请求
r = requests.post(url, data=(json.dumps(dataJson, ensure_ascii=False)).encode("utf-8"))
r = requests.post(url, data=json.dumps(dataJson))
'''
r.encoding                                 # 获取当前的编码
r.encoding = 'utf-8'                       # 设置编码
r.text                                     # 以encoding解析返回内容。字符串方式的响应体, 会自动根据响应头部的字符编码进行解码。
r.content                                  # 以字节形式（二进制）返回。字节方式的响应体, 会自动为你解码 gzip 和 deflate 压缩。
r.headers                                  # 以字典对象存储服务器响应头, 但是这个字典比较特殊, 字典键不区分大小写, 若键不存在则返回None
r.status_code                              # 响应状态码
r.raw                                      # 返回原始响应体，也就是urllib的response对象, 使用 r.raw.read()   
r.ok                                       # 查看r.ok的布尔值便可以知道是否登陆成功
*特殊方法*
r.json()                                   # Requests中内置的JSON解码器,以json形式返回,前提返回的内容确保是json格式的, 不然解析出错会抛异常
r.raise_for_status()                       # 失败请求(非200响应)抛出异常
r.headers                                  #返回字典类型,头信息
r.requests.headers                         #返回发送到服务器的头信息
r.cookies                                  #返回cookie
r.history                                  #返回重定向信息,当然可以在请求是加上allow_redirects = false 阻止重定向
'''
print(r.json())

# ======================================================================================
# 正式版本

import requests
import json
import xlrd

url = "http://zycrmqas.zeus.com:8000/ZWEB/DKHLX?sap-client=600"

dataJson = {
    "ZNB": "1000000100",
    "ZDATA": [
        {
            "MATNR_FROM": "党参破壁饮片",
            "MEINS_FROM": "罐 ",
            "MENGE": 1.0,
            "PARTNER": "1100002253",
            "ZDATE_DEAL": "20210301",
            "ZGGE_FROM": "2g×20袋",
            "ZQDCY_FROM": "劳动公园店"
        },
        {
            "MATNR_FROM": "红景天破壁饮片",
            "MEINS_FROM": "罐 ",
            "MENGE": 0.0,
            "PARTNER": "1100002253",
            "ZDATE_DEAL": "20210306",
            "ZGGE_FROM": "1g×20袋",
            "ZQDCY_FROM": "东莱街分店"
        },
        {
            "MATNR_FROM": "红景天破壁饮片",
            "MEINS_FROM": "罐 ",
            "MENGE": 1.0,
            "PARTNER": "1100002253",
            "ZDATE_DEAL": "20210311",
            "ZGGE_FROM": "1g×20袋",
            "ZQDCY_FROM": "劳动公园店"
        }
    ]
}


def read_xlsx_file(filename):
    data = xlrd.open_workbook(filename)
    table = data.sheets()[0]
    rows = table.nrows
    data = []
    for i in range(1, rows):
        values = table.row_values(i)
        data.append(
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
    return data


dataExecel = read_xlsx_file("C:/Users/Zeus/Desktop/网上客户流向3月份数据.xlsx")
# 字典中的数据都是单引号,但是标准的json需要双引号
dataJson = json.dumps(dataExecel, sort_keys=True, ensure_ascii=False, indent=4, separators=(',', ':'))
# 前面的数据只是数组,加上外面的json格式大括号
dataJson = """{""" + """\n"ZNB": "100000010",\n"ZDATA":""""" + dataJson + """\n}"""
print(dataJson)

'''
requests.get()                                # GET请求
requests.post()                               # POST请求
requests.put()                                # PUT请求
requests.delete()                             # DELETE请求
requests.head()                               # HEAD请求
requests.options()                            # OPTIONS请求
'''
# r = requests.post(url, data=(json.dumps(dataJson, ensure_ascii=False)).encode("utf-8"))
r = requests.post(url, data=dataJson.encode("utf-8"))
r = requests.post(url, data=json.dumps(dataJson))
'''
r.encoding                                 # 获取当前的编码
r.encoding = 'utf-8'                       # 设置编码
r.text                                     # 以encoding解析返回内容。字符串方式的响应体, 会自动根据响应头部的字符编码进行解码。
r.content                                  # 以字节形式（二进制）返回。字节方式的响应体, 会自动为你解码 gzip 和 deflate 压缩。
r.headers                                  # 以字典对象存储服务器响应头, 但是这个字典比较特殊, 字典键不区分大小写, 若键不存在则返回None
r.status_code                              # 响应状态码
r.raw                                      # 返回原始响应体，也就是urllib的response对象, 使用r.raw.read()   
r.ok                                       # 查看r.ok的布尔值便可以知道是否登陆成功
*特殊方法*
r.json()                                   # Requests中内置的JSON解码器,以json形式返回,前提返回的内容确保是json格式的, 不然解析出错会抛异常
r.raise_for_status()                       # 失败请求(非200响应)抛出异常
r.headers                                  #返回字典类型,头信息
r.requests.headers                         #返回发送到服务器的头信息
r.cookies                                  #返回cookie
r.history                                  #返回重定向信息,当然可以在请求是加上allow_redirects = false 阻止重定向
'''
print(r.json())

# ======================================================================================
# 成功版本

import requests
import json
import xlrd


def excel_to_json(filename, dataId):
    workbook = xlrd.open_workbook(filename)
    sheet = workbook.sheet_by_index(0)
    rows = sheet.nrows
    content = []
    for i in range(1, rows):
        values = sheet.row_values(i)
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
    dataJson = """{""" + """\n"ZNB": """"" + str(dataId) + """",\n"ZDATA":""""" + Json + """\n}"""
    return dataJson


dataExecelJson = excel_to_json("C:/Users/Zeus/Desktop/网上客户流向3月份数据.xlsx","1231000001")
print(dataExecelJson)

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
r.raw                                      # 返回原始响应体，也就是urllib的response对象, 使用r.raw.read()   
r.ok                                       # 查看r.ok的布尔值便可以知道是否登陆成功
*特殊方法*
r.json()                                   # Requests中内置的JSON解码器,以json形式返回,前提返回的内容确保是json格式的, 不然解析出错会抛异常
r.raise_for_status()                       # 失败请求(非200响应)抛出异常
r.headers                                  #返回字典类型,头信息
r.requests.headers                         #返回发送到服务器的头信息
r.cookies                                  #返回cookie
r.history                                  #返回重定向信息,当然可以在请求是加上allow_redirects = false 阻止重定向
'''
# r = requests.post(url, data=(json.dumps(dataJson, ensure_ascii=False)).encode("utf-8"))
# r = requests.post(url, data=json.dumps(dataJson))
r = requests.post("http://zycrmqas.zeus.com:8000/ZWEB/DKHLX?sap-client=600", data=dataExecelJson.encode("utf-8"))
print(r.json())
