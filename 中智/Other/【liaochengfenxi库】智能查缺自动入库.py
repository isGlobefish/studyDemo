# -*- coding: utf-8 -*-
'''
@createTool    : PyCharm-2020.2.2
@projectName   : pythonProject 
@originalAuthor: Made in win10.Sys design by deHao.Zou
@createTime    : 2020/10/15 15:43
'''
import time
import datetime
import calendar
import pymysql
import pandas as pd
from sqlalchemy import create_engine  # 连接mysql使用
from sqlalchemy.types import NVARCHAR
from termcolor import cprint
from time import strftime, gmtime
from dateutil.relativedelta import relativedelta

allStartTime = datetime.datetime.now()

YearQY = input("输入全亿年份:").zfill(4)
MonthQY = input("输入全亿月份:").zfill(2)


# Day = input("输入日数:").zfill(2)  # 前一天

# today = int(time.strftime("%d", time.localtime()))  # 本日
# if today == 1:
#     Year = str(time.strftime("%Y", time.localtime())).zfill(4)  # 本年
#     Month = str(int(time.strftime("%m", time.localtime())) - 1).zfill(2)  # 前一月
#     Day = str(calendar.monthrange(Year, Month)[1]).zfill(2)  # 前一月最后一日
# else:
#     Year = str(time.strftime("%Y", time.localtime())).zfill(4)  # 本年
#     Month = str(int(time.strftime("%m", time.localtime()))).zfill(2)  # 本月
#     Day = str(int(time.strftime("%d", time.localtime())) - 1).zfill(2)  # 前一日

# 定义上传数据库对象
def uploadSQL(uploadTable, targetDB, updateDBTable):
    '''
    uploadTable   :待上传表格
    TargetDB      ：目标数据库
    updateDBTable ：目标数据库下的表格
    '''
    try:
        # 转换数据表格式
        def mapping_df_types(convTabFormat):
            dtypedict = {}
            for i, j in zip(convTabFormat.columns, convTabFormat.dtypes):
                if "object" in str(j):
                    dtypedict.update({i: NVARCHAR(length=255)})
                if "float" in str(j):
                    dtypedict.update({i: NVARCHAR(length=255)})
                if "int" in str(j):
                    dtypedict.update({i: NVARCHAR(length=255)})
    except Exception as e:
        print("数据格式约束出错！！！", e)
    try:
        dtypedict = mapping_df_types(uploadTable)  # 转换数据格式
        engine = create_engine("mysql+pymysql://alex:123456@192.168.249.150:3306/" + str(targetDB) + "?charset=utf8")
        uploadTable.to_sql(updateDBTable, engine, dtype=dtypedict, index=False, if_exists='append')
    except Exception as e:
        print("上传数据库过程出错！！！", e)


# 定义创建表对象
def createDF():
    newDF = pd.DataFrame(columns=["公司", "日期", "店名", "销售单号", "会员卡号", "品名", "销售数量", "标准零售价", "销售金额"])
    return newDF


# 获取需要导入数据的月份
def getEveryMonth(nearestMonth, toMonth):
    monthList = []
    nearestMonth = datetime.datetime.strptime(str(nearestMonth)[0:8] + '01', "%Y-%m-%d")  # 月初1号
    while nearestMonth <= toMonth:
        dateStrZZ = nearestMonth.strftime("%Y-%m-%d")
        monthList.append(dateStrZZ)
        nearestMonth += relativedelta(months=1)
    return monthList


# 上传全亿数据
try:
    allStartTimeQY = datetime.datetime.now()
    # 读取全亿数据
    try:
        dataQY = pd.read_excel("D:/OTHER/QY" + MonthQY + ".xlsx", header=0)
    except Exception as e:
        print("读取数据失败or无数据集！！！", e)
    # 上传到【全亿数据库】
    startTimeQY1 = datetime.datetime.now()
    print("【全亿数据库】QY" + str(MonthQY) + ".xlsx入库ing......")
    connQY1 = pymysql.connect(host='192.168.249.150',  # 数据库地址
                              port=3306,  # 数据库端口
                              user='alex',  # 用户名
                              passwd='123456',  # 数据库密码
                              db='全亿数据',  # 数据库名
                              charset='utf8')  # 字符串类型
    cursorQY1 = connQY1.cursor()
    executeDelCodeQY1 = "DELETE FROM `quanyi_data` WHERE YEAR(`销售日期`)='" + YearQY + "' AND MONTH(`销售日期`)='" + MonthQY + "'"
    cursorQY1.execute(executeDelCodeQY1)
    delRowNumQY1 = cursorQY1.rowcount
    connQY1.commit()  # 提交确认
    print("【全亿数据库】开始上传ing......")
    uploadSQL(dataQY, "全亿数据", 'quanyi_data')  # 上传数据
    cursorQY1.close()  # 关闭光标
    connQY1.close()  # 关闭连接
    endTimeQY1 = datetime.datetime.now()
    cprint("【全亿数据库】全亿" + YearQY + "年" + MonthQY + "月份删除旧数据" + str(delRowNumQY1) + "行; 导入新数据" + str(
        len(dataQY)) + "行; 新增数据" + str(len(dataQY) - delRowNumQY1) + "行; 耗时：" + strftime("%H:%M:%S", gmtime(
        (endTimeQY1 - startTimeQY1).seconds)), 'magenta', attrs=['bold', 'reverse', 'blink'])
    startTimeQY2 = datetime.datetime.now()
    print("【liaochengfenxi库】QY" + str(MonthQY) + ".xlsx入库ing......")
    dfQY = createDF()  # 全亿
    dfQY["日期"] = dataQY["销售日期"]
    dfQY["店名"] = dataQY["门店名称"]
    dfQY["销售单号"] = dataQY["小票号"]
    dfQY["品名"] = dataQY["物料描述"]
    dfQY["销售数量"] = dataQY["销售数量"]
    dfQY["销售金额"] = dataQY["销售金额"]
    dfQY["公司"] = "全亿"
    print("新表匹配数据完成！！！")
    connQY2 = pymysql.connect(host='192.168.249.150',  # 数据库地址
                              port=3306,  # 数据库端口
                              user='alex',  # 用户名
                              passwd='123456',  # 数据库密码
                              db='liaochengfenxi',  # 数据库名
                              charset='utf8')  # 字符串类型
    cursorQY2 = connQY2.cursor()
    executeDelCodeQY2 = "DELETE FROM `liaocheng_sale_fact` WHERE `公司`='全亿' AND YEAR(`日期`)='" + YearQY + "' AND MONTH(`日期`)='" + MonthQY + "'"
    cursorQY2.execute(executeDelCodeQY2)
    delRowNumQY2 = cursorQY2.rowcount
    connQY2.commit()  # 提交确认
    print("【liaocheng_sale_fact库】开始上传ing......")
    uploadSQL(dfQY, 'liaochengfenxi', 'liaocheng_sale_fact')  # 上传数据
    cursorQY2.close()  # 关闭光标
    connQY2.close()  # 关闭连接
    endTimeQY2 = datetime.datetime.now()
    cprint("【liaocheng_sale_fact库】全亿" + YearQY + "年" + MonthQY + "月份删除旧数据" + str(delRowNumQY2) + "行; 导入新数据" + str(
        len(dfQY)) + "行; 新增数据" + str(len(dfQY) - delRowNumQY2) + "行; 耗时：" + strftime("%H:%M:%S", gmtime(
        (endTimeQY2 - startTimeQY2).seconds)), 'magenta', attrs=['bold', 'reverse', 'blink'])
    allEndTimeQY = datetime.datetime.now()
    print("全亿总耗时：" + strftime("%H:%M:%S", gmtime((allEndTimeQY - allStartTimeQY).seconds)))
except Exception as e:
    print("获取or上传全亿数据过程出错！！!", e)

# 上传中智数据
allStartTimeZZ = datetime.datetime.now()
getMonthStartTime = datetime.datetime.now()
print("获取中智需上传数据月份ing......")
Z0Z = pymysql.connect(
    host='192.168.249.150',
    port=3306,
    user='alex',
    passwd='123456',
    db='liaochengfenxi',
    charset='utf8'
)
cursorZZ = Z0Z.cursor()
executeZZ = "SELECT MAX(`日期`) FROM liaocheng_sale_fact WHERE `公司`='中智'"
cursorZZ.execute(executeZZ)
timeStr = cursorZZ.fetchall()
nearestTime = (str(timeStr).split(',')[0][-4:] + "-" + str(timeStr).split(',')[1] + "-" + str(timeStr).split(',')[
    2]).replace(" ", "")
Z0Z.commit()  # 提交确认
cursorZZ.close()  # 关闭光标
Z0Z.close()  # 关闭连接
nearestTimeZZ = datetime.datetime.strptime(nearestTime, "%Y-%m-%d")  # 最新日期
todayTimeZZ = datetime.datetime.now()  # 本日
missMonthZZ = getEveryMonth(nearestTimeZZ, todayTimeZZ)  # 获取需上传月份
getMonthEndTime = datetime.datetime.now()
cprint("【liaocheng_sale_fact库】中智最新日期：" + str(nearestTimeZZ) + "; 即将导入" + str(
    missMonthZZ) + "时间段内的数据; 获取上传月份耗时：" + strftime("%H:%M:%S", gmtime((getMonthEndTime - getMonthStartTime).seconds)),
       'cyan', attrs=['bold', 'reverse', 'blink'])
for iMonth in missMonthZZ:
    YearZZ = iMonth[0:4]  # 年
    MonthZZ = iMonth[5:7]  # 月
    try:
        # 删除【liaochengfenxi库】旧数据
        print("【liaocheng_sale_fact库】删除中智" + YearZZ + "年" + MonthZZ + "月份旧数据ing......")
        delStartTime = datetime.datetime.now()
        connZZ1 = pymysql.connect(host='192.168.249.150',  # 数据库地址
                                  port=3306,  # 数据库端口
                                  user='alex',  # 用户名
                                  passwd='123456',  # 数据库密码
                                  db='liaochengfenxi',  # 数据库名
                                  charset='utf8')  # 字符串类型
        cursorZZ1 = connZZ1.cursor()
        executeDelCodeZZ1 = "DELETE FROM `liaocheng_sale_fact` WHERE `公司`='中智' AND YEAR(`日期`)='" + YearZZ + "' AND MONTH(`日期`)='" + MonthZZ + "'"
        cursorZZ1.execute(executeDelCodeZZ1)
        delRowNumZZ1 = cursorZZ1.rowcount
        connZZ1.commit()  # 提交确认
        cursorZZ1.close()  # 关闭光标
        connZZ1.close()  # 关闭连接
        delEndTime = datetime.datetime.now()
        cprint(
            "【liaocheng_sale_fact库】中智" + YearZZ + "年" + MonthZZ + "月份删除旧数据" + str(delRowNumZZ1) + "行; 耗时:" + strftime(
                "%H:%M:%S", gmtime((delEndTime - delStartTime).seconds)), 'cyan', attrs=['bold', 'reverse', 'blink'])

        # 获取【dkh库】下v_sale_test
        startTimeZZ = datetime.datetime.now()
        startGetDataTime = datetime.datetime.now()
        print("【dkh库】获取中智" + YearZZ + "年" + MonthZZ + "月份数据ing......")
        connZZ2 = pymysql.connect(
            host='192.168.249.150',
            port=3306,
            user='alex',
            passwd='123456',
            db='dkh',
            charset='utf8'
        )
        cursorZZ2 = connZZ2.cursor()
        executeZZ2 = "SELECT FINALTIME, FLAG_NAME, SALENO, MEMBERCARDNO, WARENAME, WAREQTY, STDAMT, NETAMT FROM v_sale_test WHERE YEAR(FINALTIME) = '" + YearZZ + "' AND MONTH(FINALTIME) = '" + MonthZZ + "'"
        cursorZZ2.execute(executeZZ2)
        dataZZ = cursorZZ2.fetchall()
        getRowNumZZ2 = cursorZZ2.rowcount
        connZZ2.commit()  # 提交确认
        cursorZZ2.close()  # 关闭光标
        connZZ2.close()  # 关闭连接
        endGetDataTime = datetime.datetime.now()
        cprint("【dkh库】中智" + YearZZ + "年" + MonthZZ + "月份获取数据" + str(getRowNumZZ2) + "行; 耗时：" + strftime(
            "%H:%M:%S", gmtime((endGetDataTime - startGetDataTime).seconds)), 'cyan',
               attrs=['bold', 'reverse', 'blink'])
        dataZZ = pd.DataFrame(dataZZ)
        dataZZ.columns = ['FINALTIME', 'FLAG_NAME', 'SALENO', 'MEMBERCARDNO', 'WARENAME', 'WAREQTY', 'STDAMT', 'NETAMT']
        dfZZ = createDF()  # 中智
        dfZZ["日期"] = dataZZ["FINALTIME"]
        dfZZ["店名"] = dataZZ["FLAG_NAME"]
        dfZZ["销售单号"] = dataZZ["SALENO"]
        dfZZ["会员卡号"] = dataZZ["MEMBERCARDNO"]
        dfZZ["品名"] = dataZZ["WARENAME"]
        dfZZ["销售数量"] = dataZZ["WAREQTY"]
        dfZZ["标准零售价"] = dataZZ["STDAMT"]
        dfZZ["销售金额"] = dataZZ["NETAMT"]
        dfZZ["公司"] = "中智"
        print("【liaochengfenxi库】中智" + YearZZ + "年" + MonthZZ + "月份数据入库ing......")
        uploadSQL(dfZZ, 'liaochengfenxi', 'liaocheng_sale_fact')  # 上传数据
        endTimeZZ = datetime.datetime.now()
        cprint("【liaocheng_sale_fact库】中智" + YearZZ + "年" + MonthZZ + "月份删除旧数据" + str(delRowNumZZ1) + "行; 导入新数据" + str(
            len(dfZZ)) + "行; 新增数据" + str(len(dfZZ) - delRowNumZZ1) + "行; 耗时：" + strftime("%H:%M:%S", gmtime(
            (endTimeZZ - startTimeZZ).seconds)), 'cyan', attrs=['bold', 'reverse', 'blink'])
    except Exception as e:
        print("获取or上传中智数据过程出错！！!", e)
allEndTimeZZ = datetime.datetime.now()
print("中智总耗时：" + strftime("%H:%M:%S", gmtime((allEndTimeZZ - allStartTimeZZ).seconds)))
allEndTime = datetime.datetime.now()
print("总耗时：" + strftime("%H:%M:%S", gmtime((allEndTime - allStartTime).seconds)))
