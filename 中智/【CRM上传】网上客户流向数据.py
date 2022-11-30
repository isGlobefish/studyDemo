# -*- coding: utf-8 -*-
'''
@createTool    : PyCharm-2020.3.3
@projectName   : pythonProjectPy3.9
@originalAuthor: Made in win10.Sys design by deHao.Zou
@createTime    : 2020/3/23 18:39
'''
import json
import requests
import datetime
import pandas as pd
from termcolor import cprint
from time import strftime, gmtime

startTime = datetime.datetime.now()

month = input(">>>请输入上传月份: ")

if month == '0':
    # 注意！！！上传单号格式：年 + 月 + 业务伙伴编码("1100002253")后四位, 如2021032253
    sapPartnerId = input(">>>请输入更新业务伙伴单号:")


    # 更新或删除指定账号数据
    def excel_json_sap(filepath):
        print(">> 读取数据中,请稍等片刻")
        fullData = pd.read_excel(filepath, header=0)  # 读取全部数据

        startUpTime = datetime.datetime.now()

        print(">> 上传单号" + str(sapPartnerId) + "数据" + str(len(fullData)) + "行中")
        content = []
        for j in range(0, len(fullData)):
            content.append(
                (
                    {
                        "PARTNER": str(fullData.loc[j, 'PARTNER']),
                        "ZQDCY_FROM": str(fullData.loc[j, 'ZQDCY_FROM']),
                        "MATNR_FROM": str(fullData.loc[j, 'MATNR_FROM']),
                        "ZGGE_FROM": str(fullData.loc[j, 'ZGGE_FROM']),
                        "MENGE": str(fullData.loc[j, 'MENGE']),
                        "MEINS_FROM": str(fullData.loc[j, 'MEINS_FROM']),
                        "ZDATE_DEAL": str(fullData.loc[j, 'ZDATE_DEAL']),
                    }
                )

            )
        Json = json.dumps(content, sort_keys=True, ensure_ascii=False, indent=4, separators=(',', ':'))
        dataExecelJson = """{""" + """\n"ZNB": """ + """\"""" + sapPartnerId + """\"""" + """,\n"ZDATA":""" + Json + """\n}"""
        r = requests.post("http://zycrmprd01.zeus.com:8000/ZWEB/DKHLX?sap-client=800", data=dataExecelJson.encode("utf-8"), headers={"Connection":"close"})
        endUpTime = datetime.datetime.now()
        print(">> " + str(r.json()) + "\n\t单号:" + str(sapPartnerId) + " 耗时：" + strftime("%H:%M:%S", gmtime((endUpTime - startUpTime).seconds)))


    excel_json_sap("D:/FilesCenter/OnlineCustomer/DKH202103_OCData.xlsx")

else:

    # 分业务伙伴方式上传
    def excel_json_sap(filepath):
        print(">> 读取数据中,请稍等片刻")
        startSumTime = datetime.datetime.now()
        fullData = pd.read_excel(filepath, header=0)  # 读取全部数据
        fullNum = len(fullData)  # 总行数
        partnerIdList = fullData['PARTNER'].drop_duplicates().values.tolist()  # 值转化为列表形式
        partnerIdListNum = len(partnerIdList)  # 本次上传业务伙伴编码总数
        sumRows = 0

        for i, con in enumerate(partnerIdList, start=1):

            # try:
            startUpTime = datetime.datetime.now()

            sapDataId = '2021' + str(month).zfill(2) + str(con)[-4:]  # 上传单号
            partData = fullData[fullData['PARTNER'] == con]  # 获取指定业务伙伴数据
            partData = partData.reset_index(drop=True)
            sumRows += len(partData)
            print(">  " + str(i).zfill(3) + "/" + str(partnerIdListNum).zfill(3) + " 上传伙伴" + str(con) + "数据" + str(len(partData)) + "行中")

            content = []
            for j in range(0, len(partData)):
                content.append(
                    (
                        {
                            "PARTNER": str(partData.loc[j, 'PARTNER']),
                            "ZQDCY_FROM": str(partData.loc[j, 'ZQDCY_FROM']),
                            "MATNR_FROM": str(partData.loc[j, 'MATNR_FROM']),
                            "ZGGE_FROM": str(partData.loc[j, 'ZGGE_FROM']),
                            "MENGE": str(partData.loc[j, 'MENGE']),
                            "MEINS_FROM": str(partData.loc[j, 'MEINS_FROM']),
                            "ZDATE_DEAL": str(partData.loc[j, 'ZDATE_DEAL']),
                        }
                    )

                )
            # 字典中的数据都是单引号,但是标准的json需要双引号
            Json = json.dumps(content, sort_keys=True, ensure_ascii=False, indent=4, separators=(',', ':'))
            # 前面的数据只是数组,加上外面的json格式大括号, ZNB:导入SAP的流水单号
            dataExecelJson = """{""" + """\n"ZNB": """ + """\"""" + sapDataId + """\"""" + """,\n"ZDATA":""" + Json + """\n}"""
            # 可读可写, 如果不存在则创建, 如果有内容则覆盖
            # jsFile = open("D:/FilesCenter/OnlineCustomer/jsonTest" + str(i + 1) + ".json", "w+", encoding='utf-8')
            # jsFile.write(dataExecelJson)
            # jsFile.close()

            # r = requests.post(url, data=(json.dumps(dataJson, ensure_ascii=False)).encode("utf-8"))
            # r = requests.post(url, data=json.dumps(dataJson))
            # 正式环境(92.1/s)：http://zycrmprd01.zeus.com:8000/ZWEB/DKHLX?sap-client=800
            # 正式环境(92.1/s)：http://192.168.20.215:8000/ZWEB/DKHLX?sap-client=800
            # 测试环境(866.86/s)：http://zycrmqas.zeus.com:8000/ZWEB/DKHLX?sap-client=600
            r = requests.post("http://192.168.20.215:8000/ZWEB/DKHLX?sap-client=800", data=dataExecelJson.encode("utf-8"), headers={ 'Connection': 'close'})
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
            endUpTime = datetime.datetime.now()
            endSumTime = datetime.datetime.now()
            print(">> " + str(r.json()) + "\n\t单号:" + str(sapDataId) + " 耗时：" + strftime("%H:%M:%S", gmtime((endUpTime - startUpTime).seconds)))
            cprint("累计上传:" + str(sumRows) + "行; 占比:" + format(sumRows / fullNum, '.2%') + "; 累计耗时:" + strftime("%H:%M:%S", gmtime(
                (endSumTime - startSumTime).seconds)), 'cyan', attrs=['bold', 'reverse', 'blink'])
            # except Exception as e:
            #     cprint(">> 单号:" + str(sapDataId) + "; 业务伙伴:" + str(con) + "出错！！！" + str(e), 'magenta', attrs=['bold', 'reverse', 'blink'])
            #     # break  # 遇错中断
            #     pass  # 遇错跳过


    excel_json_sap("D:/FilesCenter/OnlineCustomer/DKH2021" + str(month).zfill(2) + "_OCData.xlsx")

endTime = datetime.datetime.now()
print(">>>总耗时:" + strftime("%H:%M:%S", gmtime((endTime - startTime).seconds)))
