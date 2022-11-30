import win32com

xlAscending = 1
xlSortColumns = 1
xlYes = 1

Application = win32com.client.gencache.EnsureDispatch("Excel.Application")  # 启动excel
Application.Visible = False  # 可视化
Application.DisplayAlerts = False  # 是否显示警告
wb = Application.Workbooks.Open('C:/Users/Zeus/Desktop/123.xlsx', ReadOnly=False)  # 打开excel
ws = wb.Worksheets('③老百姓')  # 选择Sheet
ws.Activate()  # 激活当前工作表



ws.Range('A4:M33').Sort(Key1=ws.Range('H1'), Order1=xlAscending, header=xlYes, Orientation=xlSortColumns)



# ws.Range(D6: D110).Sort(Key1=ws.Range('D1'), Order1=xlAscending,
#                         Key2=ws.Range('E1'), Order2=xlAscending,
#                         Key3=ws.Range('G1'), Order3=xlAscending,
#                         header=xlYes, Orientation=xlSortColumns)
#
# ws.Range(D6: D110).Sort(Key1=ws.Range('H1'), Order1=xlAscending,
#                         header=xlYes, Orientation=xlSortColumns)
