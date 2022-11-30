# -*- coding: utf-8 -*-
'''
@createTool    : PyCharm-2020.2.2
@projectName   : pythonCode
@originalAuthor: Made in win10.Sys design by deHao.Zou
@createTime    : 2020/10/12 11:39
'''
import os
import re
import time
import xlwt
import glob
import xlrd
import shutil
import pymysql
import calendar
import datetime
import win32api as ap
import pyautogui as pg
import pyperclip as cp  # 复制粘贴
from termcolor import cprint
import win32com.client as win32
from time import strftime, gmtime


def downloadToupload():
    startTime = datetime.datetime.now()  # 开始总计时
    # 解决win32模块历史运行记录影响后续代码运行问题
    clearFiles = 'C:/Users/Long/AppData/Local/Temp/gen_py/3.8/'  # 需清空文件路径
    shutil.rmtree(clearFiles)  # 递归删除文件夹
    createFiles = 'C:/Users/Long/AppData/Local/Temp/gen_py/3.8/'  # 创建文件夹路径
    os.mkdir(createFiles)  # 创建3.8空文件夹

    day = int(time.strftime("%d", time.localtime()))
    if day == 1:
        Year = str(time.strftime("%Y", time.localtime())).zfill(4)  # 本年
        Month = str(int(time.strftime("%m", time.localtime())) - 1).zfill(2)  # 前一月
        Day = str(calendar.monthrange(int(Year), int(Month))[1]).zfill(2)  # 前一月最后一日
    else:
        Year = str(time.strftime("%Y", time.localtime())).zfill(4)  # 本年
        Month = str(int(time.strftime("%m", time.localtime()))).zfill(2)  # 本月
        Day = str(int(time.strftime("%d", time.localtime())) - 1).zfill(2)  # 前一日

    # =============================================================================
    #     Year = input("输入年份:").zfill(4)
    #     Month = input("输入月份:").zfill(2)
    #     Day = input("输入日数:").zfill(2)  # 前一天
    # =============================================================================

    downloadStartTime = datetime.datetime.now()  # 开始下载数据计时

    ap.ShellExecute(0, 'open',
                    'C:\\Users\\Long\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs\\Windows PowerShell\\Windows PowerShell',
                    '', '', 1)  # 打开Windows PowerShell
    time.sleep(2)
    pg.click(267, 242)  # 点击屏幕
    inputCode = 'login-powerBI -environment China'
    passWord = 'Zeus3737'
    executeCode = "$activities = Get-PowerBIActivityEvent -StartDateTime '" + str(Year) + "-" + str(Month) + "-" + str(
        Day) + "T00:00:00' -EndDateTime '" + str(Year) + "-" + str(Month) + "-" + str(
        Day) + "T23:59:59' | ConvertFrom-Json >Z:\Hao\PowerBi_Log_To_xlsx\PowerBi_Log_To_Txt\pbidata" + str(
        Year) + "." + str(Month) + "." + str(Day) + ".txt"
    closePowerShell = 'exit'
    cp.copy(inputCode)  # 复制内容
    pg.hotkey('ctrl', 'v')  # 粘贴
    pg.press('enter')  # 回车
    time.sleep(9)  # 等待选择账号前的是
    pg.click(754, 477)  # 点击是
    pg.click(754, 477)  # 点击是
    pg.click(730, 270)  # 点击是
    time.sleep(2)
    pg.click(764, 571)  # 选择账号(第一个账号)
    pg.click(695, 300)  # 点击密码
    cp.copy(passWord)  # 复制内容
    pg.hotkey('ctrl', 'v')  # 粘贴
    pg.click(938, 404)  # 登录
    time.sleep(3)
    cp.copy(executeCode)  # 复制内容
    pg.hotkey('ctrl', 'v')  # 粘贴
    pg.press('enter')  # 回车
    time.sleep(10)  # 等待下载文件的时间
    os.chdir('Z:/Hao/PowerBi_Log_To_xlsx/')  # 定位目录根
    downLoadTxtFile = glob.glob(
        'PowerBi_Log_To_Txt/pbidata' + str(Year) + "." + str(Month) + "." + str(Day) + ".txt")  # 定位文件夹下是否已经存在下载文件
    if len(downLoadTxtFile) == 1:
        print(">  @" + str(Year) + "." + str(Month) + "." + str(Day) + "日志数据下载成功！！！")
    else:
        time.sleep(20)
        downLoadTxtFile = glob.glob(
            'PowerBi_Log_To_Txt/pbidata' + str(Year) + "." + str(Month) + "." + str(Day) + ".txt")  # 定位文件夹下是否已经存在下载文件
        if len(downLoadTxtFile) == 1:
            print(">  @" + str(Year) + "." + str(Month) + "." + str(Day) + "日志数据下载成功！！！")
        else:
            print("不存在可用文件，可能等待时间不够或其它")
    downloadEndTime = datetime.datetime.now()  # 结束下载数据计时
    print(">> 下载日志数据耗时：" + strftime("%H:%M:%S", gmtime((downloadEndTime - downloadStartTime).seconds)))

    if len(downLoadTxtFile) == 1:  # 存在日志数据执行以下程序
        # 关闭Windows PowerShell
        pg.click(267, 242)  # 点击屏幕
        cp.copy(closePowerShell)  # 复制内容
        pg.hotkey('ctrl', 'v')  # 粘贴
        pg.press('enter')  # 回车

        def uploadDB():
            # 处理数据并上传数据库
            print(">  准备数据中......")
            # 去掉txt文件里面的空白行，并保存到新的文件中
            with open("PowerBi_Log_To_Txt/pbidata" + str(Year) + "." + str(Month) + "." + str(Day) + ".txt", "r",
                      encoding='UTF-16 LE') as fr, open('PowerBI_Other/123.txt', 'w', encoding='utf-8') as fd:
                for text in fr.readlines():
                    if text.split():
                        fd.write(text)
            '''
            方法二:
            with open('output.txt', encoding='utf-8') as fp_in:
                with open('new22.txt', 'w', encoding='utf-8') as fp_out:
                    fp_out.writelines(line for i, line in enumerate(fp_in) if i != 0)
            '''
            with open('PowerBI_Other/123.txt', 'r', encoding='utf-8') as d1, open('PowerBI_Other/456.txt', 'w',
                                                                                  encoding='utf-8') as d2:
                d2.write(''.join(d1.readlines()[1:]))

            strList = []
            with open('PowerBI_Other/456.txt', 'r', encoding='utf-8') as d3, open('PowerBI_Other/789.txt', 'w',
                                                                                  encoding='utf-8') as d4:
                lines = d3.readlines()  # 读取每一行
                for line in lines:
                    strList.append(line)
                for i in range(len(strList)):
                    if strList[i][0:9] == 'UserAgent':
                        if strList[i + 1][0:9] != '         ':
                            d4.write(strList[i])
                        else:
                            if strList[i + 2][0:9] != '         ':
                                d4.write(str(strList[i]).rstrip() + str(strList[i + 1]).lstrip())
                            else:
                                d4.write(str(strList[i]).rstrip() + str(strList[i + 1]).lstrip())
                    elif strList[i][0:9] == '         ':
                        pass
                    else:
                        if strList[i][0:12] == 'CreationTime':
                            d4.write(
                                strList[i] + "Date               : " + re.sub('\s+', '', strList[i]).strip()[
                                                                       13:23] + "\n" + "Time               : " + re.sub(
                                    '\s+', '',
                                    strList[
                                        i]).strip()[
                                                                                                                 24:32] + "\n" + "DateTime           : " + re.sub(
                                    '\s+', '', strList[i]).strip()[13:23] + " " + re.sub('\s+', '', strList[i]).strip()[
                                                                                  24:32] + "\n")
                        else:
                            d4.write(strList[i])

            # 创建一个workbook对象，相当于创建一个Excel文件
            book = xlwt.Workbook(encoding='utf-8', style_compression=0)
            '''
            Workbook类初始化时有encoding和style_compression参数
            encoding:设置字符编码，一般要这样设置：w = Workbook(encoding='utf-8')，就可以在excel中输出中文了,默认是ascii。
            style_compression:表示是否压缩，不常用。
            '''
            # 创建一个sheet对象，一个sheet对象对应Excel文件中的一张表格。
            sheet = book.add_sheet('Sheet1', cell_overwrite_ok=True)
            # 其中的Output是这张表的名字,cell_overwrite_ok，表示是否可以覆盖单元格，其实是Worksheet实例化的一个参数，默认值是False

            colName = ['Id', 'RecordType', 'CreationTime', 'Date', 'Time', 'DateTime', 'Operation', 'OrganizationId',
                       'UserType', 'UserKey', 'Workload', 'UserId', 'ClientIP', 'UserAgent', 'Activity', 'ItemName',
                       'WorkSpaceName', 'DatasetName', 'ReportName', 'WorkspaceId', 'ObjectId', 'DatasetId', 'ReportId',
                       'IsSuccess', 'ReportType', 'RequestId', 'ActivityId', 'DistributionMethod', 'ConsumptionMethod',
                       'DashboardName', 'DashboardId', 'Datasets', 'SharingInformation',
                       'ExportEventStartDateTimeParameter', 'ExportEventEndDateTimeParameter']

            # 向表中添加数据标题
            for colnum, colname in zip(range(len(colName)), colName):
                sheet.write(0, colnum, colname)  # 其中的'0-行, col-列'指定表中的单元，colname是向该单元写入列名称

            # sheet.write(0, 0, 'Id')  # 其中的'0-行, 0-列'指定表中的单元，'Id'是向该单元写入的内容
            # sheet.write(0, 1, 'RecordType')
            # sheet.write(0, 2, 'CreationTime')
            # sheet.write(0, 3, 'Operation')
            # sheet.write(0, 4, 'OrganizationId')
            # sheet.write(0, 5, 'UserType')
            # sheet.write(0, 6, 'UserKey')
            # sheet.write(0, 7, 'Workload')
            # sheet.write(0, 8, 'UserId')
            # sheet.write(0, 9, 'ClientIP')
            # sheet.write(0, 10, 'UserAgent')
            # sheet.write(0, 11, 'Activity')
            # sheet.write(0, 12, 'ItemName')
            # sheet.write(0, 13, 'WorkSpaceName')
            # sheet.write(0, 14, 'DatasetName')
            # sheet.write(0, 15, 'ReportName')
            # sheet.write(0, 16, 'WorkspaceId')
            # sheet.write(0, 17, 'ObjectId')
            # sheet.write(0, 18, 'DatasetId')
            # sheet.write(0, 19, 'ReportId')
            # sheet.write(0, 20, 'IsSuccess')
            # sheet.write(0, 21, 'ReportType')
            # sheet.write(0, 22, 'RequestId')
            # sheet.write(0, 23, 'ActivityId')
            # sheet.write(0, 24, 'DistributionMethod')
            # sheet.write(0, 25, 'ConsumptionMethod')

            # 对文本内容进行多次切片得到想要的部分，然后写入指定的表格
            n = 0
            strList2 = []
            with open('PowerBI_Other/789.txt', 'r+', encoding='utf-8') as fc:
                for strText in fc.readlines():
                    strList2.append(strText)
                for applist in strList2:
                    titleName = applist.split(':', 1)[0].strip()
                    content = applist.split(':', 1)[1].strip('\n')
                    if titleName[0:2] == 'Id':
                        n = n + 1
                        sheet.write(n, 0, content)  # 往表格里写入内容
                    for locNum in range(len(colName)):
                        if colName[locNum] == titleName:
                            bNum = locNum
                            sheet.write(n, bNum, content)  # 往表格里写入内容
            # 最后，将以上操作保存到指定的Excel文件中(在此之前先判断是否存在文件，存在就删除)
            deleteFileList = os.listdir('PowerBi_Log_To_xls/')
            exitOldExcel = glob.glob(
                "PowerBi_Log_To_xls/pbidata" + str(Year) + "." + str(Month) + "." + str(Day) + ".xls")
            if len(exitOldExcel) == 1:
                print(">  已存在pbidata" + str(Year) + "." + str(Month) + "." + str(Day) + ".xls数据")
                if (len(exitOldExcel) != 0):
                    for deletefile in deleteFileList:
                        isDeleteFile = os.path.join('PowerBi_Log_To_xls/', deletefile)
                        if os.path.isfile(isDeleteFile):
                            os.remove(isDeleteFile)
                exitNewExcel = glob.glob(
                    "PowerBi_Log_To_xls/pbidata" + str(Year) + "." + str(Month) + "." + str(Day) + ".xls")
                if len(exitNewExcel) == 0:
                    book.save("PowerBi_Log_To_xls/pbidata" + str(Year) + "." + str(Month) + "." + str(Day) + ".xls")
                    print(">  覆盖旧数据完成！！！")
            else:
                book.save("PowerBi_Log_To_xls/pbidata" + str(Year) + "." + str(Month) + "." + str(Day) + ".xls")
            time.sleep(0.5)
            # Excel格式转换：.xls ---> .xlsx
            fileName = "Z:\\Hao\\PowerBi_Log_To_xlsx\\PowerBi_Log_to_xls\\pbidata" + str(Year) + "." + str(
                Month) + "." + str(Day) + ".xls"
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            wb = excel.Workbooks.Open(fileName)
            wb.SaveAs(fileName + "x", FileFormat=51)  # FileFormat = 51转为.xlsx、FileFormat = 56转为.xls
            wb.Close()
            excel.Application.Quit()
            print(">> " + str(Year) + "." + str(Month) + "." + str(Day) + ".xlsx" + "数据已准备好！！！")
            time.sleep(0.5)

            print(">>>开始把" + str(Year) + "." + str(Month) + "." + str(Day) + ".xlsx" + "数据导入数据库......")
            # 连接数据库
            try:
                database = pymysql.connect(host='192.168.249.150',  # 数据库地址
                                           port=3306,  # 数据库端口
                                           user='alex',  # 用户名
                                           passwd='123456',  # 数据库密码
                                           db='powerbi',  # 数据库名
                                           charset='utf8')  # 字符串类型
            except:
                print("连接数据库出错！！！")

            # 读取数据
            def open_Excel():
                try:
                    book = xlrd.open_workbook(
                        "Z:/Hao/PowerBi_Log_To_xlsx/PowerBi_Log_to_xls/pbidata" + str(Year) + "." + str(
                            Month) + "." + str(Day) + ".xlsx")
                except:
                    print("读取数据出错！！！")
                try:
                    sheet = book.sheet_by_name('Sheet1')
                    return sheet
                except:
                    print("读取Sheet子页出错！！！")

            # 导入数据进数据库
            def insert_Data():
                sheet = open_Excel()
                cursor = database.cursor()
                row_Num = sheet.nrows
                sql_delect = "DELETE FROM pbi_log WHERE day(Date) = '" + Day + "'and month(Date) = '" + Month + "' and year(Date) = '" + Year + "'"
                cursor.execute(sql_delect)  # SQL语句删除当月旧数据
                delRowNum = cursor.rowcount
                for i in range(1, row_Num):
                    # 第一行是标题名，对应表中的字段名所以应该从第二行开始，计算机以0开始计数，所以值是1
                    row_data = sheet.row_values(i)
                    value = (
                        row_data[0], row_data[1], row_data[2], row_data[3], row_data[4], row_data[5], row_data[6],
                        row_data[7], row_data[8], row_data[9], row_data[10], row_data[11], row_data[12], row_data[13],
                        row_data[14], row_data[15], row_data[16], row_data[17], row_data[18], row_data[19],
                        row_data[20], row_data[21], row_data[22], row_data[23], row_data[24], row_data[25],
                        row_data[26], row_data[27], row_data[28], row_data[29], row_data[30], row_data[31],
                        row_data[32], row_data[33], row_data[34])

                    # value代表的是Excel表格中的每行的数据
                    print("导入" + str(Year) + "." + str(Month) + "." + str(Day) + ".xlsx" + "第" + str(i) + "行数据ing")
                    sql = 'INSERT INTO pbi_log(Id, RecordType, CreationTime, Date, Time, DateTime, Operation, OrganizationId, UserType, UserKey, Workload, UserId, ClientIP, UserAgent, Activity, ItemName, WorkSpaceName, DatasetName, ReportName, WorkspaceId, ObjectId, DatasetId, ReportId, IsSuccess, ReportType, RequestId, ActivityId, DistributionMethod, ConsumptionMethod, DashboardName, DashboardId, Datasets,SharingInformation, ExportEventStartDateTimeParameter, ExportEventEndDateTimeParameter) VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)'
                    cursor.execute(sql, value)  # 执行sql语句
                cprint(">>>删除重复数据" + str(delRowNum) + "行;" + "新增数据" + str(row_Num - 1) + "行", 'cyan',
                       attrs=['bold', 'reverse', 'blink'])
                database.commit()
                cursor.close()  # 关闭连接

            try:
                open_Excel()
                insert_Data()
                print(">>>" + str(Year) + "." + str(Month) + "." + str(Day) + "数据导入数据库成功end！！!")
            except Exception as e:
                print(str(Year) + "." + str(Month) + "." + str(Day) + "数据导入数据库出错！！！", e)

        try:
            uploadDB()
        except Exception as e:
            print("出错！！！【删除 Path 下文件】", e)  # Path = C:\Users\Long\AppData\Local\Temp\gen_py\3.8\
        endTime = datetime.datetime.now()  # 结束总计时
        print(">>>总耗时：" + strftime("%H:%M:%S", gmtime((endTime - startTime).seconds)))
    else:
        print("不存在日志数据，请确认是否下载成功！！！")


try:
    downloadToupload()
except Exception as e:
    print("出错！！！", e)
