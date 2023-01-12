# -*- coding: utf-8 -*-
'''
@createTool    : PyCharm-2020.2.2
@projectName   : pythonProjectPy3.9 
@originalAuthor: Made in win10.Sys design by deHao.Zou
@createTime    : 2020/12/9 16:47
'''
from win32com.client import Dispatch, DispatchEx
import pythoncom
from PIL import ImageGrab, Image
import win32com
import uuid
import time


# screenArea——格式类似"A1:J10"
def excelCatchScreen(file_name, sheet_name, screen_area, save_path, img_name=False, filesname = "123456"):
    pythoncom.CoInitialize()  # excel多线程相关
    # excel = DispatchEx("Excel.Application")  # 启动excel
    excel = win32com.client.gencache.EnsureDispatch("Excel.Application")  # 启动excel
    excel.Visible = False  # 可视化
    excel.DisplayAlerts = False  # 是否显示警告
    wb = excel.Workbooks.Open(file_name)  # 打开excel
    ws = wb.Sheets(sheet_name)  # 选择sheet
    time.sleep(2)
    ws.Range(screen_area).CopyPicture()  # 复制图片区域
    time.sleep(2)
    ws.Paste()  # 粘贴 ws.Paste(ws.Range('B1'))  # 将图片移动到具体位置

    # name = str(uuid.uuid4())  # 重命名唯一值
    name = str(filesname)
    new_shape_name = name[:6]
    excel.Selection.ShapeRange.Name = new_shape_name  # 将刚刚选择的Shape重命名，避免与已有图片混淆

    ws.Shapes(new_shape_name).Copy()  # 选择图片
    time.sleep(1)
    img = ImageGrab.grabclipboard()  # 获取剪贴板的图片数据
    if not img_name:
        img_name = name + ".PNG"
    img.save(save_path + img_name)  # 保存图片
    time.sleep(2)
    wb.Close(SaveChanges=0)  # 关闭工作薄，不保存
    time.sleep(1)
    excel.Quit()  # 退出excel
    pythoncom.CoUninitialize()


if __name__ == '__main__':
    excelCatchScreen("C:/Users/Zeus/Desktop/大客户数据日报2021年1月.xlsx", "Sheet2", "A1:X9", 'C:/Users/Zeus/Desktop/')
