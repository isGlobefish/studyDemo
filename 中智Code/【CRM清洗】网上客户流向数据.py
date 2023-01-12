# -*- coding: utf-8 -*-
'''
@createTool    : PyCharm-2020.3.3
@projectName   : pythonProjectPy3.9
@originalAuthor: Made in win10.Sys design by deHao.Zou
@createTime    : 2020/3/10 14:00
'''
# --------------------
# 函数名：ZSD_RFC_001   |
# ---------------------
# 传入表：ZSD008_DR     |
# -----------------------------------------------
# 字段名        | 字段类型  | 长度   |  描述        |
# -----------------------------------------------
# PARTNER      | CHAR    |  10   |  业务伙伴      |
# ZQDCY_FROM   | CHAR    |  100  |  流向门店名称   |
# MATNR_FROM   | CHAR    |  40   |  物料描述      |
# ZGGE_FROM    | CHAR    |  20   |  物料规格      |
# MENGE        | CHAR    |  50   |  数量         |
# MEINS_FROM   | CHAR    |  10   |  单位         |
# ZDATE_DEAL   | DATS    |  8    |  交易日期      |
# ZLXID        | CHAR    |  20   |  流向ID       |
# -----------------------------------------------
# 传出参数
# ZMESSAGE	CHAR	200	消息文本
# ZTYPE  	CHAR	1	消息类型: S 成功, E 错误, W 警告, I 信息, A 中断

import os
import glob
import xlrd
import datetime
import pandas as pd
from termcolor import cprint
from time import strftime, gmtime

startTime = datetime.datetime.now()


def newDF():
    # 新建一个dataframe文件
    df = pd.DataFrame(
        columns=["MaterialID", "CustomerSystem", "MatchKey", "PARTNER", "ZQDCY_FROM", "MATNR_FROM", "ZGGE_FROM", "MENGE", "MEINS_FROM", "ZDATE_DEAL"])
    return df


if __name__ == '__main__':

    # 字典匹配业务伙伴编码
    customerIdDict = {}  # 业务伙伴编码
    customerDict = xlrd.open_workbook('D:\\FilesCenter\\OnlineCustomer\\网上客户流向业务伙伴编码对照表.xlsx')
    customerSheet = customerDict.sheet_by_name('Sheet1')
    customerRow = customerSheet.nrows
    for i in range(1, customerRow):
        customerValues = customerSheet.row_values(i)
        customerIdDict[customerValues[0]] = customerValues[3]

    # 字典匹配物料描述、规格、单位
    materialDescDict = {}  # 物料描述
    materialNormDict = {}  # 物料规格
    materialUnitDict = {}  # 物料单位
    materialDict = xlrd.open_workbook('D:\\FilesCenter\\OnlineCustomer\\商品编码对照字典(标准).xlsx')
    materialSheet = materialDict.sheet_by_name('Sheet1')
    materialRow = materialSheet.nrows
    for j in range(1, materialRow):
        materialValues = materialSheet.row_values(j)
        materialDescDict[str(materialValues[1]).replace('.0', '')] = materialValues[7]
        materialNormDict[str(materialValues[1]).replace('.0', '')] = materialValues[5]
        materialUnitDict[str(materialValues[1]).replace('.0', '')] = materialValues[6]

    month = input(">>>请输入清洗月份: ")  # 网上流向数据上传(输入本月月份)

    startDayTime = datetime.datetime.now()  # 开始计时


    # 显示当前上传目录所有文件
    def showFiles(path):
        showFileList = os.listdir(path)
        all_xls = glob.glob(path + "*.xls")
        all_xlsx = glob.glob(path + "*.xlsx")
        all_csv = glob.glob(path + "*.csv")
        print("该目录下有" + '\n' + str(showFileList) + ";" + '\n' + "其中【xls:" + str(len(all_xls)) + ", xlsx:" + str(len(all_xlsx)) + ", csv:" + str(
            len(all_csv)) + "】")


    showFiles('D:/FilesCenter/EverydayUpDB/')
    print(">> 整合数据中,请稍等片刻")
    hwdf = newDF()  # 海王
    print(" > 海王")
    # df1 = pd.read_excel("D:\\FilesCenter\\EverydayUpDB\\大客户数据日报2021年" + month + "月.xlsx", sheet_name="海王" + month + "月明细", dtype=str)
    # hwdf["MaterialID"] = df1['商品SAP编码']
    # hwdf["MatchKey"] = "海王" + "-" + df1['地区']
    # hwdf["ZQDCY_FROM"] = df1["店名/区域"]
    # hwdf["MATNR_FROM"] = hwdf.apply(lambda x: materialDescDict.setdefault(x["MaterialID"], 0), axis=1)
    # hwdf["ZGGE_FROM"] = hwdf.apply(lambda x: materialNormDict.setdefault(x["MaterialID"], 0), axis=1)
    # hwdf["MENGE"] = df1["销量"]
    # hwdf["MEINS_FROM"] = hwdf.apply(lambda x: materialUnitDict.setdefault(x["MaterialID"], 0), axis=1)
    # hwdf["ZDATE_DEAL"] = df1["过账日期"]
    # hwdf["PARTNER"] = hwdf.apply(lambda x: customerIdDict.setdefault(x["MatchKey"], 0), axis=1)
    # # hwdf["ZLXID"] = 'DKH2021' + str(month)
    # # hwdf["CustomerSystem"] = "海王"
    # for i in range(len(hwdf["ZDATE_DEAL"])):
    #     hwdf.loc[i, "ZDATE_DEAL"] = hwdf.loc[i, "ZDATE_DEAL"].replace("-", "")

    lbxdf = newDF()  # 老百姓
    print(" > 老百姓")
    df1 = pd.read_excel("D:\\FilesCenter\\EverydayUpDB\\大客户数据日报2021年" + month + "月.xlsx", sheet_name="老百姓" + month + "月明细", dtype=str)
    lbxdf["MaterialID"] = df1['商品编码']
    lbxdf["MatchKey"] = "老百姓" + "-" + df1["区域"]
    lbxdf["ZQDCY_FROM"] = df1["业务部门"]
    lbxdf["MATNR_FROM"] = lbxdf.apply(lambda x: materialDescDict.setdefault(x["MaterialID"], 0), axis=1)
    lbxdf["ZGGE_FROM"] = lbxdf.apply(lambda x: materialNormDict.setdefault(x["MaterialID"], 0), axis=1)
    lbxdf["MENGE"] = df1["数量"]
    lbxdf["MEINS_FROM"] = lbxdf.apply(lambda x: materialUnitDict.setdefault(x["MaterialID"], 0), axis=1)
    lbxdf["ZDATE_DEAL"] = df1["日期"]
    lbxdf["PARTNER"] = lbxdf.apply(lambda x: customerIdDict.setdefault(x["MatchKey"], 0), axis=1)
    # lbxdf["ZLXID"] = 'DKH2021' + str(month).zfill(2)
    # lbxdf["CustomerSystem"] = "老百姓"
    for i in range(len(lbxdf["ZDATE_DEAL"])):
        lbxdf.loc[i, "ZDATE_DEAL"] = lbxdf.loc[i, "ZDATE_DEAL"].replace("-", "")

    yfdf = newDF()  # 益丰
    print(" > 益丰")
    # df1 = pd.read_excel("D:\\FilesCenter\\EverydayUpDB\\大客户数据日报2021年" + month + "月.xlsx", sheet_name="益丰" + month + "月明细", dtype=str)
    # yfdf["MaterialID"] = df1['商品编码']
    # yfdf["MatchKey"] = "益丰" + "-" + df1["公司名"]
    # yfdf["ZQDCY_FROM"] = df1["门店名称"]

    # yfdf["MATNR_FROM"] = yfdf.apply(lambda x: materialDescDict.setdefault(x["MaterialID"], 0), axis=1)
    # yfdf["ZGGE_FROM"] = yfdf.apply(lambda x: materialNormDict.setdefault(x["MaterialID"], 0), axis=1)
    # yfdf["MENGE"] = df1["销售数量"]
    # yfdf["MEINS_FROM"] = yfdf.apply(lambda x: materialUnitDict.setdefault(x["MaterialID"], 0), axis=1)
    # yfdf["ZDATE_DEAL"] = df1["销售日期"]
    # yfdf["PARTNER"] = yfdf.apply(lambda x: customerIdDict.setdefault(x["MatchKey"], 0), axis=1)
    # # yfdf["ZLXID"] = 'DKH2021' + str(month).zfill(2)
    # # yfdf["CustomerSystem"] = "益丰"
    # for i in range(len(yfdf["ZDATE_DEAL"])):
    #     yfdf.loc[i, "ZDATE_DEAL"] = yfdf.loc[i, "ZDATE_DEAL"].replace("-", "")[0:8]

    dsldf = newDF()  # 大参林
    print(" > 大参林")
    # df1 = pd.read_excel("D:\\FilesCenter\\EverydayUpDB\\大客户数据日报2021年" + month + "月.xlsx", sheet_name="大参林" + month + "月明细", dtype=str)
    # dsldf["MaterialID"] = df1['商品编码']
    # dsldf["MatchKey"] = "大参林" + "-" + df1["调整运营区"]
    # dsldf["ZQDCY_FROM"] = df1["门店描述"]
    # dsldf["MATNR_FROM"] = dsldf.apply(lambda x: materialDescDict.setdefault(x["MaterialID"], 0), axis=1)
    # dsldf["ZGGE_FROM"] = dsldf.apply(lambda x: materialNormDict.setdefault(x["MaterialID"], 0), axis=1)
    # dsldf["MENGE"] = df1["数量"]
    # dsldf["MEINS_FROM"] = dsldf.apply(lambda x: materialUnitDict.setdefault(x["MaterialID"], 0), axis=1)
    # dsldf["ZDATE_DEAL"] = df1["销售日期"]
    # dsldf["PARTNER"] = dsldf.apply(lambda x: customerIdDict.setdefault(x["MatchKey"], 0), axis=1)
    # # dsldf["ZLXID"] = 'DKH2021' + str(month).zfill(2)
    # # dsldf["CustomerSystem"] = "大参林"
    # for i in range(len(dsldf["ZDATE_DEAL"])):
    #     dsldf.loc[i, "ZDATE_DEAL"] = dsldf.loc[i, "ZDATE_DEAL"].replace("-", "")[0:8]

    gddf = newDF()  # 国大
    print(" > 国大")
    # df1 = pd.read_excel("D:\\FilesCenter\\EverydayUpDB\\GD" + month + ".xlsx", sheet_name=0, dtype=str)
    # gddf["MaterialID"] = df1['商品编码']
    # gddf["MatchKey"] = "国大" + "-" + df1["区域名称"]
    # gddf["ZQDCY_FROM"] = df1["门店名称"]
    # gddf["MATNR_FROM"] = gddf.apply(lambda x: materialDescDict.setdefault(x["MaterialID"], 0), axis=1)
    # gddf["ZGGE_FROM"] = gddf.apply(lambda x: materialNormDict.setdefault(x["MaterialID"], 0), axis=1)
    # gddf["MENGE"] = df1["数量"]
    # gddf["MEINS_FROM"] = gddf.apply(lambda x: materialUnitDict.setdefault(x["MaterialID"], 0), axis=1)
    # gddf["ZDATE_DEAL"] = df1["销售日期"]
    # gddf["PARTNER"] = gddf.apply(lambda x: customerIdDict.setdefault(x["MatchKey"], 0), axis=1)
    # # gddf["ZLXID"] = 'DKH2021' + str(month).zfill(2)
    # # gddf["CustomerSystem"] = "国大"
    # # for i in range(len(gddf["ZQDCY_FROM"])):
    # #     if "惠州" in str(gddf.loc[i, "ZQDCY_FROM"]):
    # #         gddf.loc[i, "PARTNER"] = 1100003603  # 国药控股国大药房深圳连锁有限公司 门店门称含有惠州字眼
    # #     elif "佛山" in str(gddf.loc[i, "ZQDCY_FROM"]):
    # #         gddf.loc[i, "PARTNER"] = 1100003605  # 国药控股国大药房广州连锁有限公司 门店门称含有佛山字眼
    # #     else:
    # #         pass
    # for i in range(len(gddf["ZDATE_DEAL"])):
    #     gddf.loc[i, "ZDATE_DEAL"] = gddf.loc[i, "ZDATE_DEAL"].replace("-", "")

    sydf = newDF()  # 漱玉
    print(" > 漱玉")
    # df1 = pd.read_excel("D:\\FilesCenter\\EverydayUpDB\\SY" + month + ".xlsx", sheet_name=0, dtype=str)
    # sydf["MaterialID"] = df1['货号']
    # sydf["MatchKey"] = "漱玉" + "-" + df1["公司名称"]
    # sydf["ZQDCY_FROM"] = df1["门店名称"]
    # sydf["MATNR_FROM"] = sydf.apply(lambda x: materialDescDict.setdefault(x["MaterialID"], 0), axis=1)
    # sydf["ZGGE_FROM"] = sydf.apply(lambda x: materialNormDict.setdefault(x["MaterialID"], 0), axis=1)
    # sydf["MENGE"] = df1["数量"]
    # sydf["MEINS_FROM"] = sydf.apply(lambda x: materialUnitDict.setdefault(x["MaterialID"], 0), axis=1)
    # sydf["ZDATE_DEAL"] = df1['销售日期']
    # sydf["PARTNER"] = sydf.apply(lambda x: customerIdDict.setdefault(x["MatchKey"], 0), axis=1)
    # # sydf["ZLXID"] = 'DKH2021' + str(month).zfill(2)
    # sydf["CustomerSystem"] = df1["实际区域"]
    # for i in range(len(sydf["ZDATE_DEAL"])):
    #     if "济南" in str(sydf.loc[i, "CustomerSystem"]):
    #         sydf.loc[i, "PARTNER"] = 1100002206  # 漱玉平民大药房连锁股份有限公司 按片区 济南划分
    #     elif "莱芜" in str(sydf.loc[i, "CustomerSystem"]) or "漱玉健康" in str(sydf.loc[i, "CustomerSystem"]):
    #         sydf.loc[i, "PARTNER"] = 1100002209  # 莱芜属于淄博 漱玉健康属于莱芜也就是属于淄博
    #     else:
    #         pass
    #     sydf.loc[i, "ZDATE_DEAL"] = sydf.loc[i, "ZDATE_DEAL"].replace("-", "")

    qydf = newDF()  # 全亿
    print(" > 全亿")
    # df1 = pd.read_excel("D:\\FilesCenter\\EverydayUpDB\\QY" + month + ".xlsx", sheet_name=0, dtype=str)
    # qydf["MaterialID"] = df1['物料编号']
    # qydf["MatchKey"] = "全亿" + "-" + df1["公司名称"]
    # qydf["ZQDCY_FROM"] = df1["门店名称"]
    # qydf["MATNR_FROM"] = qydf.apply(lambda x: materialDescDict.setdefault(x["MaterialID"], 0), axis=1)
    # qydf["ZGGE_FROM"] = qydf.apply(lambda x: materialNormDict.setdefault(x["MaterialID"], 0), axis=1)
    # qydf["MENGE"] = df1["销售数量"]
    # qydf["MEINS_FROM"] = qydf.apply(lambda x: materialUnitDict.setdefault(x["MaterialID"], 0), axis=1)
    # qydf["ZDATE_DEAL"] = df1['销售日期']
    # qydf["PARTNER"] = qydf.apply(lambda x: customerIdDict.setdefault(x["MatchKey"], 0), axis=1)
    # # qydf["ZLXID"] = 'DKH2021' + str(month).zfill(2)
    # # qydf["CustomerSystem"] = "全亿"
    # for i in range(len(qydf["ZDATE_DEAL"])):
    #     qydf.loc[i, "ZDATE_DEAL"] = qydf.loc[i, "ZDATE_DEAL"].replace("-", "")[0:8]

    gjdf = newDF()  # 高济
    print(" > 高济")
    # df1 = pd.read_excel("D:\\FilesCenter\\EverydayUpDB\\GJ" + month + ".xlsx", sheet_name=0, dtype=str)
    # gjdf["MaterialID"] = df1['商品编码']
    # gjdf["MatchKey"] = "高济" + "-" + df1["企业名称"]
    # gjdf["ZQDCY_FROM"] = df1["门店名称"]
    # gjdf["MATNR_FROM"] = gjdf.apply(lambda x: materialDescDict.setdefault(x["MaterialID"], 0), axis=1)
    # gjdf["ZGGE_FROM"] = gjdf.apply(lambda x: materialNormDict.setdefault(x["MaterialID"], 0), axis=1)
    # gjdf["MENGE"] = df1["销售数量"]
    # gjdf["MEINS_FROM"] = gjdf.apply(lambda x: materialUnitDict.setdefault(x["MaterialID"], 0), axis=1)
    # gjdf["ZDATE_DEAL"] = df1["业务日期"]
    # gjdf["PARTNER"] = gjdf.apply(lambda x: customerIdDict.setdefault(x["MatchKey"], 0), axis=1)
    # # gjdf["ZLXID"] = 'DKH2021' + str(month).zfill(2)
    # # gjdf["CustomerSystem"] = "高济"
    # for i in range(len(gjdf["ZDATE_DEAL"])):
    #     gjdf.loc[i, "ZDATE_DEAL"] = gjdf.loc[i, "ZDATE_DEAL"].replace("/", "")

    print(">> 存储数据中, 请稍等片刻")
    alldf = hwdf.append([lbxdf, dsldf, yfdf, gddf, sydf, qydf, gjdf])  # 整合所有数据于一个表格下
    # alldf["MaterialID"] = pd.to_numeric(alldf["MaterialID"])
    alldf["MENGE"] = pd.to_numeric(alldf["MENGE"])
    alldf["ZDATE_DEAL"] = alldf["ZDATE_DEAL"].astype(int)
    alldf["PARTNER"] = alldf["PARTNER"].replace("-", 1100000000) # 电商
    # --------------------------------------------------------------------------------------------
    # 去掉空格 去掉字母
    # import pandas as pd
    # import re
    # data = pd.read_excel('C:/Users/Zeus/Desktop/456.xlsx', header=0, sheet_name=0)
    # str = '桂平市老百姓浔州路大药房 连锁店 '
    # print(re.sub('\xa0', '',data['sfa_desc'][0].replace(' ', '')))


    # alldf = alldf.drop(columns=["MaterialID", "CustomerSystem", "MatchKey"])
    # alldf = alldf[~alldf["PARTNER"].isin('-')]
    alldf.to_excel('D:/FilesCenter/OnlineCustomer/DKH2021' + month.zfill(2) + '_OCData.xlsx', index=False)

    endTime = datetime.datetime.now()
    print(">>>总耗时：" + strftime("%H:%M:%S", gmtime((endTime - startTime).seconds)))
    cprint(">>>2021" + month.zfill(2) + "数据导出成功！ >> 注意核查数据匹配情况", 'magenta', attrs=['bold', 'reverse', 'blink'])
