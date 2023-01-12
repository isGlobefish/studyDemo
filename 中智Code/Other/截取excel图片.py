# -*- coding: utf-8 -*-
'''
@createTool    : PyCharm-2020.3.2
@projectName   : pythonProject
@originalAuthor: Made in win10.Sys design by deHao.Zou
@createTime    : 2021/01/14 20:51
'''

import os
import win32com
from win32com.client import constants as c  # 旨在直接使用VBA常数

# current_address = os.path.abspath('.')
current_address = os.path.abspath('.')
excel_address = os.path.join(current_address, "郭云峰.xlsx")
print(current_address)
print(excel_address)
xl_app = win32com.client.gencache.EnsureDispatch("Excel.Application")  # 若想引用常数的话使用此法调用Excel
xl_app.Visible = False  # 是否显示Excel文件
wb = xl_app.Workbooks.Open(excel_address)
sht = wb.Worksheets(1)
sht.Name = "示例"

sht.Range("A1").Value = "小试牛刀"
wb.Save()
wb.Close()

# from win32com.client import Dispatch, DispatchEx
import win32com
from win32com.client import constants as c  # 旨在直接使用VBA常数
import pythoncom
from PIL import ImageGrab, Image
import uuid


# screenArea——格式类似"A1:J10"
def excelCatchScreen(file_name, sheet_name, screen_area, save_path, img_name=False):
    pythoncom.CoInitialize()  # excel多线程相关
    excel = win32com.client.gencache.EnsureDispatch("Excel.Application")  # 启动excel
    excel.Visible = True  # 可视化
    excel.DisplayAlerts = False  # 是否显示警告
    wb = excel.Workbooks.Open(file_name)  # 打开excel
    ws = wb.Sheets(sheet_name)  # 选择Sheet
    ws.Range(screen_area).CopyPicture()  # 复制图片区域
    ws.Paste()  # 粘贴 ws.Paste(ws.Range('B1'))  # 将图片移动到具体位置

    # name = str(uuid.uuid4())  # 重命名唯一值
    name = "拜访数据可视化"
    new_shape_name = name[:6]
    excel.Selection.ShapeRange.Name = new_shape_name  # 将刚刚选择的Shape重命名, 避免与已有图片混淆

    ws.Shapes(new_shape_name).Copy()  # 选择图片
    img = ImageGrab.grabclipboard()  # 获取剪贴板的图片数据
    if not img_name:
        img_name = name + ".PNG"
    img.save(save_path + img_name)  # 保存图片
    # wb.Save()
    wb.Close(SaveChanges=0)  # 关闭工作薄，不保存
    excel.Quit()  # 退出excel
    pythoncom.CoUninitialize()


excelCatchScreen('D:/DataCenter/VisImage/Files/贾宝玉.xls', "Sheet1", "A1:M16", 'D:/DataCenter/VisImage/Image/')
