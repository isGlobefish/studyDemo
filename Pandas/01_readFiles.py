'''
逝者如斯夫, 不舍昼夜 -- 孔夫子
@Auhor    : Dohoo Zou
Project   : gitCode
FileName  : 01_readFiles.py
IDE       : PyCharm
CreateTime: 2022-11-26 18:11:26
'''
# 加载模块
import pandas as pd

fpath = '/Users/dohozou/Desktop/Code/gitCode/Pandas/dataFiles/'

# 读取csv文件数据
data_csv = pd.read_csv(fpath + 'pandas_csv.csv', header=0)
# 查看数据前五行
data_csv.head()
# 查看数据后五行
data_csv.tail()
# 查看数据形状（行和列）
data_csv.shape
# 查看数据列名
data_csv.columns
# 查看数据索引
data_csv.index
# 查看数据每一列的类型
data_csv.dtypes

# 读取txt文件数据
data_txt = pd.read_csv(fpath + 'pandas_txt.txt', header=None, sep=',', names=['A', 'B', 'C', 'D'])

# 读取xlsx文件数据
data_xlsx = pd.read_excel(fpath + 'pandas_xlsx.xlsx', sheet_name=0)

# 读取数据库中的数据
import pymysql

data = pymysql.connect(
    host='localhost',
    port=3306,
    user='root',
    password='88888888',
    database='???',
    charset='utf8'
)

















