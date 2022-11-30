# -*- coding: utf-8 -*-
'''
@createTool    : PyCharm-2020.2.2
@projectName   : pythonProjectPy3.9
@originalAuthor: Made in win10.Sys design by deHao.Zou
@createTime    : 2020/12/8 16:18
'''
import re
import os
import glob
import xlrd
import pymysql
import calendar
import datetime
import openpyxl
import pandas as pd
from termcolor import cprint
from time import strftime, gmtime
from sqlalchemy import create_engine  # 连接mysql使用
from sqlalchemy.types import NVARCHAR

startTime = datetime.datetime.now()


def mapping_df_types(df):
    dtypedict = {}
    for i, j in zip(df.columns, df.dtypes):
        if "object" in str(j):
            dtypedict.update({i: NVARCHAR(length=255)})
        if "float" in str(j):
            dtypedict.update({i: NVARCHAR(length=255)})
        if "int" in str(j):
            dtypedict.update({i: NVARCHAR(length=255)})


def get_excel(wei_zhi):
    all_excel = glob.glob(wei_zhi + "*.xlsx")  # 匹配所有xlsx文件
    print("该目录下有" + str(len(all_excel)) + "个excel文件：")
    if (len(all_excel) == 0):
        return 0
    else:
        for i in range(len(all_excel)):
            print(' > ' + re.split(' |/|\\\\', str(all_excel[i]))[-1])
        return all_excel


def write_excel_xlsx(path, sheet_name, value):  # 导出excel表格
    index = len(value)
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = sheet_name
    for i in range(0, index):
        for j in range(0, len(value[i])):
            sheet.cell(row=i + 1, column=j + 1, value=str(value[i][j]))
    workbook.save(path)
    print(">  xlsx格式表格写入数据成功！")


def newDF():
    # 新建一个dataframe文件
    df = pd.DataFrame(
        columns=['desc_1', 'DKH_materiel_id', 'materiel_desc', 'norms', 'date', 'sfa_desc', 'sfa_id', 'amount', 'UnitPrice', 'sales_Money',
                 'bz_UnitPrice', 'bz_sales_Money', 'customer', 'materiel_alias', 'dept_5', 'sfa_client_desc', 'state', 'client_id', 'client_desc',
                 'client_alias'])
    return df


if __name__ == '__main__':
    # 字典匹配
    priceDict = {}  # 匹配单价
    materielDict = {}  # 匹配商品简称
    dictionary = xlrd.open_workbook('D:\\FilesCenter\\DKH-BottomTable\\商品编码对照字典.xlsx')  # 载入字典
    sheet = dictionary.sheet_by_name('Sheet1')
    row = sheet.nrows
    for i in range(1, row):
        values = sheet.row_values(i)
        priceDict[str(values[1]).replace('.0', '')] = values[8]
        materielDict[str(values[1]).replace('.0', '')] = values[6]

    gjcomDict = {}  # 高济平台
    gjcompany = xlrd.open_workbook('C:\\Users\\Zeus\\Desktop\\autoSend\\大客户\\目标\\大客户_数据源.xlsx')  # 载入字典
    gjsheet = gjcompany.sheet_by_name('高济目标公司')
    gjrow = gjsheet.nrows
    for irow in range(1, gjrow):
        gjvalues = gjsheet.row_values(irow)
        gjcomDict[gjvalues[0]] = gjvalues[2]


    # k 为 0 时 导入大数据源数据; k 为 本月月份时 导入other文件夹数据

    k = input(">>>请输入月份：")

    if k == "0":  # 上传每月大改数据

        uploadType = input(">>>输入上传类型【A(全部上传)、S(选择上传)】：")

        if uploadType == 'A' or uploadType == 'a':  # 上传全部数据
            startAllMonthTime = datetime.datetime.now()
            startMonthTime = datetime.datetime.now()  # 开始计时
            all_excel = get_excel("D:\\FilesCenter\\大客户数据源\\")  # 获取路径下全部xls文件
            alldf = newDF()
            for i in range(len(all_excel)):
                # 逐一读取EXCEL文件
                print("开始合成 :", re.split(' |/|\\\\', str(all_excel[i]))[-1])
                df = pd.read_excel(all_excel[i], sheet_name=0, dtype=str)  # df-通过表单索引来指定读取的表单，第一个文件第二个表格
                alldf = alldf.append([df])  # 合并在到df数据表中
            alldf['date'] = pd.to_datetime(alldf['date'], format='%Y/%m/%d').dt.date
            allRowsNum = len(alldf)  # 所有数据条数
            print(">  转换数据格式")
            dtypedict = mapping_df_types(alldf)
            endMonthTime = datetime.datetime.now()
            print(">> 整合数据" + str(allRowsNum) + "行, 耗时:" + strftime("%H:%M:%S", gmtime((endMonthTime - startMonthTime).seconds)))

            startDBTime = datetime.datetime.now()
            print(">  上传MySQLing: LinuxDB -> dkh -> dkhfact")
            # 设置mysql连接引擎
            engine = create_engine('mysql+pymysql://root:Powerbi#1217@192.168.20.241:3306/dkh', encoding='utf-8',
                                   echo=False, pool_size=100, max_overflow=10, pool_timeout=100, pool_recycle=7200)
            alldf.to_sql('dkhfact', engine, dtype=dtypedict, index=False, if_exists='replace')
            endDBTime = datetime.datetime.now()
            print(">> 全部上传耗时:" + strftime("%H:%M:%S", gmtime((endDBTime - startDBTime).seconds)))

            # 导出新增门店
            startOutTime = datetime.datetime.now()
            dfMerge = df.drop_duplicates(subset=['dept_5', 'customer'], keep='first')  # 去重，保存第一条重复数据
            dfMerge = dfMerge.reset_index(drop=True)
            dfMerge["orgDataColumn"] = dfMerge["dept_5"].map(str) + dfMerge["customer"]  # 列拼接合并
            refExcel = pd.read_excel("D:\\FilesCenter\\DKH-BottomTable\\DKH对照表.xlsx", sheet_name="SFA_Hierarchy", dtype=str)
            refDataSplit = pd.DataFrame((x.split('-') for x in refExcel['OUT_Dept_1']), index=refExcel.index, columns=['leftSplit', 'rightSplit'])
            refSplitMergeTab = pd.merge(refExcel, refDataSplit, right_index=True, left_index=True)
            refSplitMergeTab["refDataColumn"] = refSplitMergeTab["OUT_SFA_bianma"].map(str) + refSplitMergeTab['rightSplit']
            toList = refSplitMergeTab['refDataColumn'].drop_duplicates().values.tolist()  # 值转化为列表形式
            dfCreate = newDF()
            for i in range(len(dfMerge)):
                if dfMerge.at[i, 'orgDataColumn'] not in toList:
                    dfCreate = dfCreate.append(dfMerge.loc[[i]])
            # 增加客户英文前缀
            dfCreate = dfCreate.reset_index(drop=True)
            customerCname = ["老百姓", "海王", "益丰", "大参林", "国大", "全亿", "漱玉", "高济"]
            customerEname = ["LBX", "HW", "YF", "DSL", "GD", "QY", "SY", "GJ"]
            for i in range(len(dfCreate)):
                for c, e in zip(customerCname, customerEname):
                    if dfCreate.loc[i, "customer"] == c:
                        dfCreate.loc[i, "material_alias"] = e + "-" + c
                        dfCreate.loc[i, "DKH_material_id"] = e + "-" + str(dfCreate.loc[i, "desc_1"])
                    else:
                        pass
            dfCreate.to_excel("D:\\FilesCenter\\DKH-BottomTable\\【Linux】新增门店.xlsx")
            endOutTime = datetime.datetime.now()
            print(">> 导出新增门店耗时：" + strftime("%H:%M:%S", gmtime((endOutTime - startOutTime).seconds)))
            endAllMonthTime = datetime.datetime.now()
            print(">> 全部上传总耗时：" + strftime("%H:%M:%S", gmtime((endAllMonthTime - startAllMonthTime).seconds)))

        else:  # 选择上传大改部分

            Year = input(">>>输入待上传数据的年份：")
            Month = input(">>>输入待上传数据的月份：").zfill(2)
            firstDay = '01'  # 前一个月第一天
            lastDay = str(calendar.monthrange(int(Year), int(Month))[1]).zfill(2)  # 前一月最后一日


            # 选择数据上传

            def dkhOrginalFiles(path):

                startSelectTime = datetime.datetime.now()
                startMakeTime = datetime.datetime.now()

                # 显示目录下所有的excel文件
                dkhFileList = os.listdir(path)
                all_Xls = glob.glob(path + "*.xls")
                all_Xlsx = glob.glob(path + "*.xlsx")
                all_Csv = glob.glob(path + "*.csv")
                print("该目录下有" + '\n' + str(dkhFileList) + ";" + '\n' + "其中【xls:" + str(len(all_Xls)) + ", xlsx:" + str(
                    len(all_Xlsx)) + ", csv:" + str(len(all_Csv)) + "】")

                # 选择符合时间内的数据归一表格
                dfCreate = newDF()
                for dkhindex, dkhfile in enumerate(dkhFileList, start=1):  # 遍历所有文件
                    fileName = os.path.splitext(dkhfile)[0]  # 获取文件名
                    fileType = os.path.splitext(dkhfile)[1]  # 获取文件扩展名
                    fileFullPath = path + fileName + fileType  # 文件完整路径
                    print(str(dkhindex).zfill(2) + '/' + str(len(dkhFileList)) + " 数据读取进程：" + dkhfile)
                    togetData = pd.read_excel(fileFullPath, header=0)
                    togetData['date'] = pd.to_datetime(togetData['date'])
                    selectData = togetData[(togetData['date'] >= pd.to_datetime(Year + Month + firstDay)) & (
                            togetData['date'] <= pd.to_datetime(Year + Month + lastDay))]
                    selectData = selectData.reset_index(drop=True)
                    dfCreate = dfCreate.append([selectData])  # 合并数据
                dfCreate['date'] = pd.to_datetime(dfCreate['date'], format='%Y/%m/%d').dt.date
                selRowsNum = len(dfCreate)  # 选择时间的上传数据行数

                # 上传前删除选择时间段内历史数据
                conn = pymysql.connect(
                    host='192.168.20.241',
                    port=3306,
                    user='root',
                    passwd='Powerbi#1217',
                    db='dkh',
                    charset='utf8'
                )
                cursor = conn.cursor()  # 获取游标
                sql_delect = "DELETE FROM dkhfact WHERE MONTH(date) = '" + Month + "' and YEAR(date) = '" + Year + "'"
                cursor.execute(sql_delect)  # SQL语句删除当月旧数据
                delRowNum = cursor.rowcount  # 删除旧数据条数
                conn.commit()  # 提交确认
                cursor.close()  # 关闭光标
                conn.close()  # 关闭连接
                endMakeTime = datetime.datetime.now()
                cprint(Year + '.' + Month + "删除旧数据" + str(delRowNum) + "行; 即将导入新数据" + str(selRowsNum) + "行; 新增数据" + str(
                    selRowsNum - delRowNum) + "行; 删除及整合数据耗时:" + strftime("%H:%M:%S", gmtime((endMakeTime - startMakeTime).seconds)), 'cyan',
                       attrs=['bold', 'reverse', 'blink'])

                # 连接mysql引擎
                startUpTime = datetime.datetime.now()
                dtypedict = mapping_df_types(dfCreate)
                print(">  上传MySQLing: LinuxDB -> dkh -> dkhfact")
                engine = create_engine('mysql+pymysql://root:Powerbi#1217@192.168.20.241:3306/dkh', encoding='utf-8',
                                       echo=False, pool_size=100, max_overflow=10, pool_timeout=100, pool_recycle=7200)
                dfCreate.to_sql('dkhfact', engine, dtype=dtypedict, index=False, if_exists='append')
                endUpTime = datetime.datetime.now()
                print(">> 选择上传耗时：" + strftime("%H:%M:%S", gmtime((endUpTime - startUpTime).seconds)))

                # 导出新增门店
                startOutTime = datetime.datetime.now()
                dfMerge = dfCreate.drop_duplicates(subset=['dept_5', 'customer'], keep='first')  # 去重，保存第一条重复数据
                dfMerge = dfMerge.reset_index(drop=True)
                dfMerge["orgDataColumn"] = dfMerge["dept_5"].map(str) + dfMerge["customer"]  # 列拼接合并
                refExcel = pd.read_excel("D:\\FilesCenter\\DKH-BottomTable\\DKH对照表.xlsx", sheet_name="SFA_Hierarchy", dtype=str)
                refDataSplit = pd.DataFrame((x.split('-') for x in refExcel['OUT_Dept_1']), index=refExcel.index, columns=['leftSplit', 'rightSplit'])
                refSplitMergeTab = pd.merge(refExcel, refDataSplit, right_index=True, left_index=True)
                refSplitMergeTab["refDataColumn"] = refSplitMergeTab["OUT_SFA_bianma"].map(str) + refSplitMergeTab['rightSplit']
                toList = refSplitMergeTab['refDataColumn'].drop_duplicates().values.tolist()  # 值转化为列表形式
                dfCreateNew = newDF()
                for i in range(len(dfMerge)):
                    if dfMerge.at[i, 'orgDataColumn'] not in toList:
                        dfCreateNew = dfCreateNew.append(dfMerge.loc[[i]])
                # 增加客户英文前缀
                dfCreateNew = dfCreateNew.reset_index(drop=True)
                customerCname = ["老百姓", "海王", "益丰", "大参林", "国大", "全亿", "漱玉", "高济"]
                customerEname = ["LBX", "HW", "YF", "DSL", "GD", "QY", "SY", "GJ"]
                for i in range(len(dfCreateNew)):
                    for c, e in zip(customerCname, customerEname):
                        if dfCreateNew.loc[i, "customer"] == c:
                            dfCreateNew.loc[i, "material_alias"] = e + "-" + c
                            dfCreateNew.loc[i, "DKH_material_id"] = e + "-" + str(dfCreateNew.loc[i, "desc_1"])
                        else:
                            pass
                dfCreateNew.to_excel("D:\\FilesCenter\\DKH-BottomTable\\【Linux】新增门店.xlsx")
                endOutTime = datetime.datetime.now()
                print(">> 导出新增门店耗时：" + strftime("%H:%M:%S", gmtime((endOutTime - startOutTime).seconds)))
                endSelectTime = datetime.datetime.now()
                print(">> 选择上传总耗时:" + strftime("%H:%M:%S", gmtime((endSelectTime - startSelectTime).seconds)))


            dkhOrginalFiles('D:\\FilesCenter\\大客户数据源\\')

    # 每日大客户数据上传(输入k数字是本月月份)
    else:
        startDayTime = datetime.datetime.now()  # 开始计时


        # 显示当前目录所有文件

        def showFiles(path):
            showFileList = os.listdir(path)
            all_Xls = glob.glob(path + "*.xls")
            all_Xlsx = glob.glob(path + "*.xlsx")
            all_Csv = glob.glob(path + "*.csv")
            print("该目录下有" + '\n' + str(showFileList) + ";" + '\n' + "其中【xls:" + str(len(all_Xls)) + ", xlsx:" + str(
                len(all_Xlsx)) + ", csv:" + str(len(all_Csv)) + "】")


        showFiles('D:/FilesCenter/EverydayUpDB/')
        month = int(k)
        print(">> 整合数据中,请稍等片刻")
        hwdf = newDF()  # 海王
        print(" > 海王")
        df1 = pd.read_excel("D:\\FilesCenter\\EverydayUpDB\\大客户数据日报2021年" + str(month) + "月.xlsx", sheet_name="海王" + str(month) + "月明细", dtype=str)
        hwdf['DKH_materiel_id'] = df1['商品SAP编码']
        hwdf['materiel_desc'] = df1['商品名称']
        hwdf['norms'] = df1['规格']
        hwdf['dept_5'] = df1['店号/区域ID']
        hwdf['sfa_id'] = df1['店号/区域ID']
        hwdf['sfa_desc'] = df1['店名/区域']
        hwdf['amount'] = df1['销量']
        hwdf['date'] = df1['过账日期']
        hwdf['desc_1'] = df1['地区']
        hwdf['UnitPrice'] = df1['单价']
        hwdf['sales_Money'] = df1['金额']
        hwdf['state'] = df1['省份']
        hwdf['bz_UnitPrice'] = hwdf.apply(lambda x: priceDict.setdefault(x['DKH_materiel_id'], 0), axis=1)
        hwdf['materiel_alias'] = hwdf.apply(lambda x: materielDict.setdefault(x['DKH_materiel_id'], 0), axis=1)
        hwdf['bz_sales_Money'] = hwdf['bz_UnitPrice'].map(float) * hwdf['amount'].map(float)
        hwdf['customer'] = "海王"

        lbxdf = newDF()  # 老百姓
        print(" > 老百姓")
        df1 = pd.read_excel("D:\\FilesCenter\\EverydayUpDB\\大客户数据日报2021年" + str(month) + "月.xlsx", sheet_name="老百姓" + str(month) + "月明细", dtype=str)
        lbxdf['DKH_materiel_id'] = df1['商品编码']
        lbxdf['materiel_desc'] = df1['商品名称']
        lbxdf['norms'] = df1['规格']
        lbxdf['dept_5'] = df1['业务部门']
        lbxdf['sfa_desc'] = df1['业务部门']
        lbxdf['amount'] = df1['数量']
        lbxdf['date'] = df1['日期']
        lbxdf['desc_1'] = df1['区域']
        lbxdf['sfa_client_desc'] = df1['区域目标']
        lbxdf['UnitPrice'] = df1['单价']
        lbxdf['sales_Money'] = df1['金额']
        lbxdf['bz_UnitPrice'] = lbxdf.apply(lambda x: priceDict.setdefault(x['DKH_materiel_id'], 0), axis=1)
        lbxdf['materiel_alias'] = lbxdf.apply(lambda x: materielDict.setdefault(x['DKH_materiel_id'], 0), axis=1)
        lbxdf['bz_sales_Money'] = lbxdf['bz_UnitPrice'].map(float) * lbxdf['amount'].map(float)
        lbxdf['customer'] = "老百姓"

        yfdf = newDF()  # 益丰
        print(" > 益丰")
        df1 = pd.read_excel("D:\\FilesCenter\\EverydayUpDB\\大客户数据日报2021年" + str(month) + "月.xlsx", sheet_name="益丰" + str(month) + "月明细", dtype=str)
        yfdf['DKH_materiel_id'] = df1['商品编码']
        yfdf['materiel_desc'] = df1['通用名']
        yfdf['norms'] = df1['规格']
        yfdf['dept_5'] = df1['门店名称']
        yfdf['sfa_desc'] = df1['门店名称']
        yfdf['amount'] = df1['销售数量']
        yfdf['date'] = df1['销售日期']
        yfdf['desc_1'] = df1['公司名']
        yfdf['UnitPrice'] = df1['单价']
        yfdf['sales_Money'] = df1['金额']
        yfdf['bz_UnitPrice'] = yfdf.apply(lambda x: priceDict.setdefault(x['DKH_materiel_id'], 0), axis=1)
        yfdf['materiel_alias'] = yfdf.apply(lambda x: materielDict.setdefault(x['DKH_materiel_id'], 0), axis=1)
        yfdf['bz_sales_Money'] = yfdf['bz_UnitPrice'].map(float) * yfdf['amount'].map(float)
        yfdf['customer'] = "益丰"

        dsldf = newDF()  # 大参林
        print(" > 大参林")
        df1 = pd.read_excel("D:\\FilesCenter\\EverydayUpDB\\大客户数据日报2021年" + str(month) + "月.xlsx", sheet_name="大参林" + str(month) + "月明细", dtype=str)
        dsldf['DKH_materiel_id'] = df1['商品编码']
        dsldf['materiel_desc'] = df1['商品名称']
        dsldf['norms'] = df1['商品规格']
        dsldf['dept_5'] = df1['门店编码']
        dsldf['sfa_id'] = df1['门店编码']
        dsldf['sfa_desc'] = df1['门店描述']
        dsldf['amount'] = df1['数量']
        dsldf['date'] = df1['销售日期']
        dsldf['desc_1'] = df1['调整运营区']
        dsldf['sfa_client_desc'] = df1['大区']
        dsldf['state'] = df1['省区']
        dsldf['client_id'] = df1['营运区']
        dsldf['UnitPrice'] = df1['单价']
        dsldf['sales_Money'] = df1['金额']
        dsldf['bz_UnitPrice'] = dsldf.apply(lambda x: priceDict.setdefault(x['DKH_materiel_id'], 0), axis=1)
        dsldf['materiel_alias'] = dsldf.apply(lambda x: materielDict.setdefault(x['DKH_materiel_id'], 0), axis=1)
        dsldf['bz_sales_Money'] = dsldf['bz_UnitPrice'].map(float) * dsldf['amount'].map(float)
        dsldf['customer'] = "大参林"

        gddf = newDF()  # 国大
        print(" > 国大")
        df1 = pd.read_excel("D:\\FilesCenter\\EverydayUpDB\\GD" + str(month) + ".xlsx", sheet_name=0, dtype=str)
        gddf['DKH_materiel_id'] = df1['商品编码']
        gddf['materiel_desc'] = df1['商品名称']
        gddf['norms'] = df1['规格']
        gddf['dept_5'] = df1['门店编码']
        gddf['sfa_id'] = df1['门店编码']
        gddf['sfa_desc'] = df1['门店名称']
        gddf['amount'] = df1['数量']
        gddf['date'] = df1['销售日期']
        gddf['desc_1'] = df1['区域名称']
        gddf['client_desc'] = df1['区域']
        gddf['client_alias'] = df1['区域名称']
        gddf['UnitPrice'] = gddf.apply(lambda x: priceDict.setdefault(x['DKH_materiel_id'], 0), axis=1)
        gddf['sales_Money'] = gddf['UnitPrice'].map(float) * gddf['amount'].map(float)
        gddf['bz_UnitPrice'] = gddf.apply(lambda x: priceDict.setdefault(x['DKH_materiel_id'], 0), axis=1)
        gddf['materiel_alias'] = gddf.apply(lambda x: materielDict.setdefault(x['DKH_materiel_id'], 0), axis=1)
        gddf['bz_sales_Money'] = gddf['bz_UnitPrice'].map(float) * gddf['amount'].map(float)
        gddf['customer'] = "国大"

        sydf = newDF()  # 漱玉
        print(" > 漱玉")
        df1 = pd.read_excel("D:\\FilesCenter\\EverydayUpDB\\SY" + str(month) + ".xlsx", sheet_name=0, dtype=str)
        sydf['DKH_materiel_id'] = df1['货号']
        sydf['materiel_desc'] = df1['品名']
        sydf['norms'] = df1['规格']
        sydf['dept_5'] = df1['门店名称']
        sydf['sfa_desc'] = df1['门店名称']
        sydf['amount'] = df1['数量']
        sydf['date'] = df1['销售日期']
        sydf['desc_1'] = df1['实际区域']
        sydf['client_desc'] = df1['公司编码']
        sydf['client_alias'] = df1['公司名称']
        sydf['sales_Money'] = df1['零售总额']
        sydf['UnitPrice'] = df1['零售总额'].map(float) / df1['数量'].map(float)
        sydf['bz_UnitPrice'] = sydf.apply(lambda x: priceDict.setdefault(x['DKH_materiel_id'], 0), axis=1)
        sydf['materiel_alias'] = sydf.apply(lambda x: materielDict.setdefault(x['DKH_materiel_id'], 0), axis=1)
        sydf['bz_sales_Money'] = sydf['bz_UnitPrice'].map(float) * sydf['amount'].map(float)
        sydf['customer'] = "漱玉"

        qydf = newDF()  # 全亿
        print(" > 全亿")
        df1 = pd.read_excel("D:\\FilesCenter\\EverydayUpDB\\QY" + str(month) + ".xlsx", sheet_name=0, dtype=str)
        qydf['DKH_materiel_id'] = df1['物料编号']
        qydf['materiel_desc'] = df1['物料描述']
        qydf['norms'] = df1['规格']
        qydf['dept_5'] = df1['门店名称']
        qydf['sfa_id'] = df1['门店代码']
        qydf['sfa_desc'] = df1['门店名称']
        qydf['date'] = df1['销售日期']
        qydf['desc_1'] = df1['公司名称']
        qydf['client_desc'] = df1['公司代码']
        qydf['client_alias'] = df1['公司名称']
        qydf['amount'] = df1['销售数量']
        qydf['sales_Money'] = df1['销售金额']
        qydf['UnitPrice'] = df1['销售金额'].map(float) / df1['销售数量'].map(float)
        qydf['bz_UnitPrice'] = qydf.apply(lambda x: priceDict.setdefault(x['DKH_materiel_id'], 0), axis=1)
        qydf['materiel_alias'] = qydf.apply(lambda x: materielDict.setdefault(x['DKH_materiel_id'], 0), axis=1)
        qydf['bz_sales_Money'] = qydf['bz_UnitPrice'].map(float) * qydf['amount'].map(float)
        qydf['customer'] = "全亿"

        gjdf = newDF()  # 高济
        print(" > 高济")
        # xl = pd.ExcelFile("D:\\FilesCenter\\EverydayUpDB\\GJ"+str(month)+".xlsx")
        # res = len(xl.sheet_names)
        # for i in range(res-1):
        # gjdf=newdf()
        df1 = pd.read_excel("D:\\FilesCenter\\EverydayUpDB\\GJ" + str(month) + ".xlsx", sheet_name=0, dtype=str)
        gjdf['DKH_materiel_id'] = df1['商品编码']
        gjdf['materiel_desc'] = df1['商品名称']
        gjdf['norms'] = df1['规格']
        gjdf['dept_5'] = df1['门店名称']
        gjdf['sfa_id'] = df1['门店编码']
        gjdf['sfa_desc'] = df1['门店名称']
        gjdf['amount'] = df1['销售数量']
        gjdf['date'] = df1['业务日期']
        gjdf['desc_1'] = df1['企业名称']
        gjdf['client_desc'] = df1['企业编码']
        gjdf['client_alias'] = df1['企业名称']
        gjdf['UnitPrice'] = gjdf.apply(lambda x: priceDict.setdefault(x['DKH_materiel_id'], 0), axis=1)
        gjdf['sales_Money'] = gjdf['UnitPrice'].map(float) * gjdf['amount'].map(float)
        gjdf['bz_UnitPrice'] = gjdf.apply(lambda x: priceDict.setdefault(x['DKH_materiel_id'], 0), axis=1)
        gjdf['materiel_alias'] = gjdf.apply(lambda x: materielDict.setdefault(x['DKH_materiel_id'], 0), axis=1)
        gjdf['bz_sales_Money'] = gjdf['bz_UnitPrice'].map(float) * gjdf['amount'].map(float)
        gjdf['sfa_client_desc'] = gjdf.apply(lambda x: gjcomDict.setdefault(x['desc_1'], ''), axis=1)
        gjdf['customer'] = "高济"

        ordf = newDF()  # 其他
        columns = ['desc_1', 'DKH_materiel_id', 'materiel_desc', 'norms', 'date', 'sfa_desc', 'sfa_id', 'amount', 'UnitPrice', 'sales_Money',
                   'bz_UnitPrice', 'bz_sales_Money', 'customer', 'materiel_alias', 'dept_5', 'sfa_client_desc', 'state', 'client_id', 'client_desc', 'client_alias']
        print(" > 其他")
        df1 = pd.read_excel("D:\\FilesCenter\\EverydayUpDB\\OR" + str(month) + ".xlsx", sheet_name=0, dtype=str)
        for conn in columns:
            ordf[conn] = df1[conn]

        print(">> 整合完毕, 请稍等片刻")
        alldf = hwdf.append([lbxdf, dsldf, yfdf, gddf, sydf, qydf, gjdf, ordf])  # 整合所有数据于一个表格下
        alldf['date'] = pd.to_datetime(alldf['date'], format='%Y/%m/%d').dt.date
        alldfRowsNum = len(alldf)  # 整合数据行数

        # 上传前删除历史数据(按月)
        print('>> 删除历史数据中,请稍等片刻')
        conn = pymysql.connect(
            host='192.168.20.241',
            port=3306,
            user='root',
            passwd='Powerbi#1217',
            db='dkh',
            charset='utf8'
        )
        cursor = conn.cursor()  # 获取游标
        # 每年初注意修改
        sql_delect = "DELETE FROM dkhfact WHERE MONTH(date) = '" + str(month).zfill(2) + "' and YEAR(date) = '2021'"
        cursor.execute(sql_delect)  # SQL语句删除当月旧数据
        delRowNum = cursor.rowcount  # 删除旧数据条数
        conn.commit()  # 提交确认
        cursor.close()  # 关闭光标
        conn.close()  # 关闭连接
        dtypedict = mapping_df_types(alldf)
        endDayTime = datetime.datetime.now()
        cprint(str(month) + "月份删除旧数据" + str(delRowNum) + "行; 即将导入新数据" + str(alldfRowsNum) + "行; 新增数据" + str(
            alldfRowsNum - delRowNum) + "行; 整合数据耗时:" + strftime("%H:%M:%S", gmtime((endDayTime - startDayTime).seconds)), 'cyan',
               attrs=['bold', 'reverse', 'blink'])

        startUploadDBTime = datetime.datetime.now()
        print(" > 上传MySQLing: LinuxDB -> dkh -> dkhfact")
        # 连接mysql引擎
        engine = create_engine('mysql+pymysql://root:Powerbi#1217@192.168.20.241:3306/dkh', encoding='utf-8',
                               echo=False, pool_size=100, max_overflow=10, pool_timeout=100, pool_recycle=7200)
        alldf.to_sql('dkhfact', engine, dtype=dtypedict, index=False, if_exists='append')
        endUploadDBTime = datetime.datetime.now()
        print(">> 上传数据库耗时:" + strftime("%H:%M:%S", gmtime((endUploadDBTime - startUploadDBTime).seconds)))

        # 导出新增门店
        startOutDayTime = datetime.datetime.now()
        dfMerge = alldf.drop_duplicates(subset=['dept_5', 'customer'], keep='first')  # 去重，保存第一条重复数据
        dfMerge = dfMerge.reset_index(drop=True)
        dfMerge["orgDataColumn"] = dfMerge["dept_5"].map(str) + dfMerge["customer"]  # 列拼接合并，合一
        refExcel = pd.read_excel("D:\\FilesCenter\\DKH-BottomTable\\DKH对照表.xlsx", sheet_name="SFA_Hierarchy", dtype=str)
        refDataSplit = pd.DataFrame((x.split('-') for x in refExcel['OUT_Dept_1']), index=refExcel.index, columns=['leftSplit', 'rightSplit'])
        refSplitMergeTab = pd.merge(refExcel, refDataSplit, right_index=True, left_index=True)
        refSplitMergeTab["refDataColumn"] = refSplitMergeTab["OUT_SFA_bianma"].map(str) + refSplitMergeTab['rightSplit']
        toList = refSplitMergeTab['refDataColumn'].drop_duplicates().values.tolist()  # 值转化为列表形式
        dfCreate = newDF()
        for i in range(len(dfMerge)):
            if dfMerge.at[i, 'orgDataColumn'] not in toList:
                dfCreate = dfCreate.append(dfMerge.loc[[i]])
        # 增加客户英文前缀
        dfCreate = dfCreate.reset_index(drop=True)

        customerCname = ['老百姓', '海王', '益丰', '大参林', '国大', '全亿', '漱玉', '高济', '湖南千金', '湖南和盛', '湖南怀仁', '吉林亚泰', '江西华氏', '正和祥', '健之佳', '一心堂', '诚民天济', '广东九州通', '河南张仲景', '四川太极', '益生天济', '柳州桂中', '国大湖北']
        customerEname = ['LBX', 'HW', 'YF', 'DSL', 'GD', 'QY', 'SY', 'GJ', 'HNQJ', 'HNHS', 'HNHR', 'JLYT', 'JXHS', 'ZHS', 'JZJ', 'YXT', 'CMTJ', 'GDJZT', 'HNZZJ', 'SCTJ', 'YSTJ', 'LZGZ', 'GDHB']
        for i in range(len(dfCreate)):
            for c, e in zip(customerCname, customerEname):
                if dfCreate.loc[i, "customer"] == c:
                    dfCreate.loc[i, "material_alias"] = e + "-" + c
                    dfCreate.loc[i, "DKH_material_id"] = e + "-" + str(dfCreate.loc[i, "desc_1"])
                else:
                    pass
        dfCreate.to_excel("D:\\FilesCenter\\DKH-BottomTable\\【Linux】新增门店.xlsx")
        endOutDayTime = datetime.datetime.now()
        print(">> 导出新增门店耗时：" + strftime("%H:%M:%S", gmtime((endOutDayTime - startOutDayTime).seconds)))

    endTime = datetime.datetime.now()
    print(">>>总耗时：" + strftime("%H:%M:%S", gmtime((endTime - startTime).seconds)))
    cprint(">>>新增门店导出成功！ >> 记得发给秀清姐哦 > ", 'magenta', attrs=['bold', 'reverse', 'blink'])
