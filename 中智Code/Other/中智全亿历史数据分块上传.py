import math
import pandas as pd
import glob
import datetime
import openpyxl
import pandas as pd
from xlrd import xldate_as_tuple
from sqlalchemy import create_engine  # 连接mysql使用
from sqlalchemy.types import Integer, NVARCHAR, Float
import pymysql
import os


# 上传数据库
def uploadSQL(uploadTable):
    try:
        def mapping_df_types(conversionFormat):
            dtypedict = {}
            for i, j in zip(conversionFormat.columns, conversionFormat.dtypes):
                if "object" in str(j):
                    dtypedict.update({i: NVARCHAR(length=255)})
                if "float" in str(j):
                    dtypedict.update({i: NVARCHAR(length=255)})
                if "int" in str(j):
                    dtypedict.update({i: NVARCHAR(length=255)})

        dtypedict = mapping_df_types(uploadTable)  # 转换数据格式
        engine = create_engine('mysql+pymysql://alex:123456@192.168.249.150:3306/123?charset=utf8')
        uploadTable.to_sql('liaocheng_sale_fact', engine, dtype=dtypedict, index=False, if_exists='append')
    except Exception as e:
        print("上传数据库过程出错！！！", e)


# 创建表对象
def createDF():
    newDF = pd.DataFrame(columns=["公司", "日期", "店名", "销售单号", "会员卡号", "品名", "销售数量", "标准零售价", "销售金额"])
    return newDF


# 上传全亿数据
try:
    connQY = pymysql.connect(
        # 获取全亿数据库下quanyi_data
        host='192.168.249.150',
        port=3306,
        user='alex',
        passwd='123456',
        db='全亿数据',
        charset='utf8'
    )

    cursorQY = connQY.cursor()
    executeQY = "SELECT 销售日期, 门店名称, 小票号, 物料描述, 销售数量, 销售金额 FROM quanyi_data"
    cursorQY.execute(executeQY)
    dataQY = cursorQY.fetchall()
    connQY.commit()
    cursorQY.close()
    connQY.close()

    dataQY = pd.DataFrame(dataQY)
    dataQY.columns = ['销售日期', '门店名称', '小票号', '物料描述', '销售数量', '销售金额']

    # 全亿数据
    dfQY = createDF()
    dfQY["日期"] = dataQY["销售日期"]
    dfQY["店名"] = dataQY["门店名称"]
    dfQY["销售单号"] = dataQY["小票号"]
    dfQY["品名"] = dataQY["物料描述"]
    dfQY["销售数量"] = dataQY["销售数量"]
    dfQY["销售金额"] = dataQY["销售金额"]
    dfQY["公司"] = "全亿"
    print("开始上传数据ing......")
    startTimeQY = datetime.datetime.now()
    uploadSQL(dfQY)
    endTimeQY = datetime.datetime.now()
    print("上传成功！！！")
    print("全亿上传耗时：" + str((endTimeQY - startTimeQY).seconds) + "秒")
except Exception as e:
    print("获取全亿数据过程出错！！!", e)



# 上传中智数据
try:
    dataZZ = pd.read_csv('G:/dataZZ2020-10-19.csv', encoding='utf-8', header=0)
    # 获取中智数据库下quanyi_data
    dfZZ = createDF()  # 中智Code
    dfZZ["日期"] = dataZZ["FINALTIME"]
    dfZZ["店名"] = dataZZ["FLAG_NAME"]
    dfZZ["销售单号"] = dataZZ["SALENO"]
    dfZZ["会员卡号"] = dataZZ["MEMBERCARDNO"]
    dfZZ["品名"] = dataZZ["WARENAME"]
    dfZZ["销售数量"] = dataZZ["WAREQTY"]
    dfZZ["标准零售价"] = dataZZ["STDAMT"]
    dfZZ["销售金额"] = dataZZ["NETAMT"]
    dfZZ["公司"] = "中智Code"
    maxRow = 1000000  # 每次上传数据量
    excelNum = math.ceil(len(dfZZ) / maxRow)  # 需要上传次数
except Exception as e:
    print("获取中智数据过程出错！！！", e)


# 分存数据
# savePath = 'G:/BigDateNew/'  # 储存路径
# saveAllStartTime = datetime.datetime.now()
# for i in range(excelNum):
#     saveStartTime = datetime.datetime.now()
#     dfZZ[(maxRow * i):(maxRow * (i + 1))].to_excel(
#         savePath + "dataZZ" + str(maxRow * i) + "-" + str(maxRow * (i + 1)) + ".xlsx", encoding='utf-8', index=False)
#     saveEndTime = datetime.datetime.now()
#     print("分存进度：" + str(i + 1) + " / " + str(excelNum) + "; 耗时:" + str((saveEndTime - saveStartTime).seconds) + "秒")
# saveAllEndTime = datetime.datetime.now()
# print("分存数据总耗时：" + str((saveAllEndTime - saveAllStartTime).seconds) + "秒")

# # 将分存数据上传数据库
# filesList = os.listdir(savePath)
# allStartTime = datetime.datetime.now()
# countUploadFiles = 0
# print("数据导入MySQLing: alex -> 123 -> sale_fact")
# for file in filesList:
#     startTime = datetime.datetime.now()
#     countUploadFiles += 1
#     fileName = os.path.split(file)[0]  # 获取文件名
#     fileType = os.path.split(file)[1]  # 获取文件格式
#     openFile = pd.read_excel(savePath + fileName + fileType, header=0)
#     uploadSQL(openFile)
#     endTime = datetime.datetime.now()
#     print("进度情况：" + str(countUploadFiles) + " / " + str(excelNum) + "; 耗时：" + str((endTime - startTime).seconds) + "秒")
# allEndTime = datetime.datetime.now()
# print("分存上传总耗时：" + str((allEndTime - allStartTime).seconds) + "秒")


# 直接分块上传
try:
    print("开始上传数据ing......")
    uploadAllStartTime = datetime.datetime.now()
    for i in range(excelNum):
        sectionStartTime = datetime.datetime.now()
        uploadSQL(dfZZ[(maxRow * i):(maxRow * (i + 1))])
        sectionEndTime = datetime.datetime.now()
        print("上传进度：" + str(i + 1) + " / " + str(excelNum) + "; 耗时：" + str(
            (sectionEndTime - sectionStartTime).seconds) + "秒")
    uploadAllEndTime = datetime.datetime.now()
    print("中智数据上传成功！！！")
    print("直接分块上传总耗时：" + str((uploadAllEndTime - uploadAllStartTime).seconds) + "秒")
except Exception as e:
    print("分块上传数据过程出错！！！", e)
