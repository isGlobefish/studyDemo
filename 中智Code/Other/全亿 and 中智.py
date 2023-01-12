# -*- coding: utf-8 -*-
'''
@createTool    : PyCharm-2020.2.2
@projectName   : pythonCode 
@originalAuthor: Made in win10.Sys design by deHao.Zou
@createTime    : 2020/10/13 14:42
'''
import glob
import datetime
import openpyxl
import pandas as pd
from xlrd import xldate_as_tuple
from sqlalchemy import create_engine  # 连接mysql使用
from sqlalchemy.types import Integer, NVARCHAR, Float
import pymysql


def mapping_df_types(df):
    dtypedict = {}
    for i, j in zip(df.columns, df.dtypes):
        if "object" in str(j):
            dtypedict.update({i: NVARCHAR(length=255)})
        if "float" in str(j):
            dtypedict.update({i: NVARCHAR(length=255)})
        if "int" in str(j):
            dtypedict.update({i: NVARCHAR(length=255)})


def createDF():
    newDF = pd.DataFrame(columns=["公司", "日期", "店名", "销售单号", "会员卡号", "品名", "销售数量", "标准零售价", "销售金额"])
    return newDF


# 获取全亿数据库下quanyi_data
connQY = pymysql.connect(
    host='192.168.249.150',
    port=3306,
    user='alex',
    passwd='123456',
    db='全亿数据',
    charset='utf8'
)
cursorQY = connQY.cursor()
executeQY = "SELECT * FROM quanyi_data"
cursorQY.execute(executeQY)
dataQY = cursorQY.fetchall()
connQY.commit()
cursorQY.close()
connQY.close()

dataQY = pd.DataFrame(dataQY)
dataQY.columns = ["集团", "小票号", "小票行号", "批号", "销售日期", "公司代码", "公司名称", "门店代码", "门店名称", "物料编号",
                  "物料描述", "规格", "基本计量单位", "计量单位文本", "生产厂家", "销售数量", "销售金额"]

# dataQY.to_csv('G:/dataQY.csv', encoding='utf-8', header=0, index=False)
# # 读取csv文件"

dataZZ = pd.read_csv('G:/v_sale_test.csv', encoding='utf-8', header=0)

# # 获取dkh数据库下v_sale_test
# connZZ = pymysql.connect(
#     host='192.168.249.150',
#     port=3306,
#     user='alex',
#     passwd='123456',
#     db='dkh',
#     charset='utf8'
# )
# cursorZZ = connZZ.cursor()
# exceuteZZ = "SELECT FINALTIME, FLAG_NAME, SALENO, MEMBERCARDNO, WARENAME, WAREQTY, STDAMT, NETAMT FROM v_sale_test"
# cursorZZ.execute(exceuteZZ)
# dataZZ = cursorZZ.fetchall()
# connZZ.commit()
# cursorZZ.close()
# connZZ.close()
#
# dataZZ = pd.DataFrame(dataZZ, index=False)
# dataZZ.columns = ["FLAG_BS", "FLAG_NO", "FLAG_NAME", "ORDER_START_TIME", "VENCUSNO", "VENCUSNAME", "OLSHOPID",
#                   "COMMNAME", "MEMBERCARDNO", "SALER", "WAREQTY", "WAREID", "WARECODE", "WARENAME", "TIMES", "STDAMT",
#                   "STALLNO", "SALENO", "ROWNO", "SALETAX", "PURTAX", "PURPRICE", "NETAMT", "PURSUM", "WAREQTY_ZJ",
#                   "MINQTY", "STDTOMIN", "NETPRICE", "MINPRICE", "MAKENO", "GROUPID", "DISTYPE", "DISRATE", "BUSNO",
#                   "BATID", "ACCDATE", "INVALIDATE", "FACTORYID", "FACTORYID1", "WARESPEC", "WAREUNIT", "SALE_RATE",
#                   "FINALTIME", "WAREGENERALNAME", "CARDHOLDER"]

dfQY = createDF()  # 全亿
dfQY["日期"] = dataQY["销售日期"]
dfQY["店名"] = dataQY["门店名称"]
dfQY["销售单号"] = dataQY["小票号"]
dfQY["品名"] = dataQY["物料描述"]
dfQY["销售数量"] = dataQY["销售数量"]
dfQY["销售金额"] = dataQY["销售金额"]
dfQY["公司"] = "全亿"
# '销售日期', '门店名称', '小票号', '物料描述', '销售数量', '销售金额'

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
# FINALTIME, FLAG_NAME, SALENO, MEMBERCARDNO, WARENAME, WAREQTY, STDAMT, NETAMT

mergeTable = dfQY.append([dfZZ])  # 合并表格


mergeTable = pd.read_csv('G:/mergeTable.csv', encoding='utf-8', header=0)

# 上传数据库
dtypedict = mapping_df_types(mergeTable)  # 转换数据格式
print("数据导入MySQLing: alex -> dkh -> liaochengfenxi")
engine = create_engine('mysql+pymysql://alex:123456@192.168.249.150:3306/liaochengfenxi?charset=utf8')
mergeTable.to_sql('sale_fact', engine, dtype=dtypedict, index=False, if_exists='replace')






# INSERT INTO quanyi_data(`集团`,`小票号`,`小票行号`,`批号`,`销售日期`,`公司代码`,`公司名称`,`门店代码`,`门店名称`,`物料编号`,`物料描述`,`规格`,`基本计量单位`,`计量单位文本`,`生产厂家`,`销售数量`,`销售金额`) VALUES(1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17)
# dataZZ.columns = ["FLAG_BS", "FLAG_NO", "FLAG_NAME", "ORDER_START_TIME", "VENCUSNO", "VENCUSNAME", "OLSHOPID",
#                   "COMMNAME", "MEMBERCARDNO", "SALER", "WAREQTY", "WAREID", "WARECODE", "WARENAME", "TIMES",
#                   "STDAMT", "STALLNO", "SALENO", "ROWNO", "SALETAX", "PURTAX", "PURPRICE", "NETAMT", "PURSUM",
#                   "WAREQTY_ZJ", "MINQTY", "STDTOMIN", "NETPRICE", "MINPRICE", "MAKENO", "GROUPID", "DISTYPE",
#                   "DISRATE", "BUSNO", "BATID", "ACCDATE", "INVALIDATE", "FACTORYID", "FACTORYID1", "WARESPEC",
#                   "WAREUNIT", "SALE_RATE", "FINALTIME", "WAREGENERALNAME", "CARDHOLDER"]




# dayPathQY = 'C:/Users/Long/Desktop/123/'
# filesList = os.listdir(dayPathQY)
# for file in filesList:
#     dataQY = pd.read_excel(dayPathQY + file, header=0)

'''
# 读取数据
def readExcel():
    try:
        book = xlrd.open_workbook("D:/OTHER/QY" + str(Month) + ".xlsx")
    except:
        print("读取数据出错！！！")
    try:
        # sheet = book.sheet_by_name('Sheet0')
        sheet = book.sheet_by_index(0)
        return sheet
    except:
        print("读取Sheet子页出错！！！")


# 连接数据库
try:
    database = pymysql.connect(host='192.168.249.150',  # 数据库地址
                               port=3306,  # 数据库端口
                               user='alex',  # 用户名
                               passwd='123456',  # 数据库密码
                               db='123',  # 数据库名
                               charset='utf8')  # 字符串类型
except:
    print("连接数据库出错！！！")
sheet = readExcel()
cursor = database.cursor()
rowNums = sheet.nrows
executeDelCode = "DELETE FROM quanyi_data WHERE month(`销售日期`) = '" + Month + "' and year(`销售日期`) = '" + Year + "'"
cursor.execute(executeDelCode)  # SQL语句删除当月旧数据
delRowNum = cursor.rowcount
print("导入QY" + str(Month) + ".xlsx数据ing......")
for i in range(1, rowNums):
    # 第一行是标题名，对应表中的字段名所以应该从第二行开始，计算机以0开始计数，所以值是1
    row_data = sheet.row_values(i)
    value = (
        row_data[0], row_data[1], row_data[2], row_data[3], row_data[4], row_data[5], row_data[6],
        row_data[7], row_data[8], row_data[9], row_data[10], row_data[11], row_data[12], row_data[13],
        row_data[14], row_data[15], row_data[16])

    # value代表的是Excel表格中的每行的数据
    sql = 'INSERT INTO quanyi_data(`集团`,`小票号`,`小票行号`,`批号`,`销售日期`,`公司代码`,`公司名称`,`门店代码`,`门店名称`,`物料编号`,`物料描述`,`规格`,`基本计量单位`,`计量单位文本`,`生产厂家`,`销售数量`,`销售金额`) VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)'
    cursor.execute(sql, value)  # 执行sql语句
print(">>>删除重复数据" + str(delRowNum) + "行;" + "新增数据" + str(rowNums - 1) + "行")
database.commit()
cursor.close()  # 关闭连接
'''
