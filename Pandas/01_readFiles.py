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

# 方法一：读取数据库中的数据
import pymysql

conn = pymysql.connect(
    host='localhost',
    port=3306,
    user='root',
    password='88888888',
    database='my_test',
    charset='utf8mb4'
)

cursor = conn.cursor()

getData = cursor.execute('select * from day20')
conn.commit()

allData = cursor.fetchall()

cursor.close()
conn.close()

# 方法二：读取数据库中的数据
from sqlalchemy import create_engine
import pandas as pd

MYSQL_HOST = 'localhost'
MYSQL_PORT = '3306'
MYSQL_USER = 'root'
MYSQL_PASSWORD = '88888888'
MYSQL_DB = 'my_test'

engine = create_engine('mysql+pymysql://%s:%s@%s:%s/%s?charset=utf8'
                       % (MYSQL_USER, MYSQL_PASSWORD, MYSQL_HOST, MYSQL_PORT, MYSQL_DB))

sql = 'select * from day20'

df = pd.read_sql(sql, engine)

# 获取设置值
pd.get_option

# print输出结果行列对齐
pd.set_option('display.unicode.ambiguous_as_wide', True)
pd.set_option('display.unicode.east_asian_width', True)
# 显示所有列
# pd.set_option('display.max_columns', None)
# 显示所有行
# pd.set_option('display.max_rows', None)
# 不换行显示
# pd.set_option('display.width', 1000)
# 显示精度
# pd.set_option('display.precision', 15)
print(df)

# 一、数据结构
# 1. Pandas的数据结构Series
import pandas as pd
import numpy as np

s1 = pd.Series([1, 'a', 3.14, 7])
s1.index
s1.values

# 创建具有标签索引的Series
s2 = pd.Series([1, 'a', 3.14, 7], index=['a', 'b', 'c', 'd'])
s2.index
s2.values

# 使用字典创建Series
dict = {'a': 100, 'b': 200, 'c': 300}
s3 = pd.Series(dict)
print(s3['a'])
print(s3['a', 'b'])
print(type(s3['a']))

# 2. Pandas的数据结构Dataframe
# 根据多个字典序列创建dataframe
data = {
    'one': ['a', 'b', 'c', 'd'],
    'two': [1, 2, 3, 4],
    'three': [3.14, 3.141, 3.1415, 3.14159]
}

df = pd.DataFrame(data)
df.index
df.columns
df.dtypes
# 查询一列
df['one']
# 查询多列
df[['one', 'two']]
type(df['one'])  # Series
type(df[['one', 'two']])  # dataframe
# 查询一行，结果是一个pd.Series
df.loc[1]
# 查询多行
df.loc[1:3]

# 二、数据查询
'''
1. df.loc方法，根据行、列标签查询
2. df.iloc方法，根据行、列的数字位置查询
3. df.where
4. df.query
.loc既能查询，又能覆盖写入
'''
import pandas as pd

fpath = '/Users/dohozou/Desktop/Code/gitCode/Pandas/dataFiles/'
weather = pd.read_excel(fpath + 'weather_xlsx.xlsx', sheet_name=0)
weather.head()
# 设置索引为日期列，方便日期筛选
weather.set_index('日期', inplace=True)
weather.head()
weather.index
# 替换温度后面的文本
weather.loc[:, '最低温'] = weather['最低温'].str.replace('"C', '').astype('int32')
weather.loc[:, '最高温'] = weather['最高温'].str.replace('"C', '').astype('int32')
weather.head()
weather.dtypes

# 2.1 使用单个label值查询数据
weather.loc['2023-01-01', "最低温"]  # 得到单个值
weather.loc['2023-01-01', ["最低温", "最高温"]]  # 得到单个series
weather.loc[['2023-01-01', '2023-01-02', '2023-01-03'], "最低温"]  # 得到series
weather.loc[['2023-01-01', '2023-01-02', '2023-01-03'], ["最低温", "最高温"]]  # 得到dataframe

# 2.2 使用数值区间进行范围查询
weather.loc['2023-01-01':'2023-01-05', '最低温']
# 列index按区间
weather.loc['2023-01-01', '最低温':'级别']
# 行和列都按区间查询
weather.loc['2023-01-01':'2023-01-05', '最低温':'级别']

# 2.3 使用条件表达式查询
weather.loc[weather['最低温'] < -30, :]
# 复杂查询使用&
weather.loc[(weather['最低温'] < -30) & (weather['最高温'] > 50) & (weather['级别'] == 1)]

# 2.4 使用函数查询
weather.loc[lambda df: (df['最低温'] < -30) & (df['最高温'] > 50), :]


def query_weather(df):
    return df.index.astype('str').str.startswith('2023-02')


weather.loc[query_weather, :]

# 三、新增数据列
'''
1. 直接赋值
2. df.apply方法
3. df.assign方
4. 按条件选择分组分别赋值
'''
import pandas as pd

fpath = '/Users/dohozou/Desktop/Code/gitCode/Pandas/dataFiles/'
weather = pd.read_excel(fpath + 'weather_xlsx.xlsx', sheet_name=0)
weather.head()
# 替换温度后面的文本
weather.loc[:, '最低温'] = weather['最低温'].str.replace('"C', '').astype('int32')
weather.loc[:, '最高温'] = weather['最高温'].str.replace('"C', '').astype('int32')
weather.head()

# 3.1. 新增温差列
weather.loc[:, '温差'] = weather['最高温'] - weather['最低温']


# 3.2. df.apply方法新增列
def get_wendu_type(x):
    if x['最低温'] < -20:
        return '低温'
    if x['最高温'] > 50:
        return '高温'


# 注意需要设置axis=1，这是series的index是colums
weather.loc[:, '温度类型'] = weather.apply(get_wendu_type, axis=1)
weather['温度类型'].value_counts()

# 3.3. df.assign方法
weather.assign(
    huashi_low=lambda x: x['最低温'] * 9 / 5 + 32,
    huashi_high=lambda x: x['最高温'] * 9 / 5 + 32
)

# 3.4. 按条件选择分组分别赋值
weather['温差max'] = ''
weather.loc[weather['最高温'] - weather['最低温'] > 100, '温差max'] = "温差大"
weather.loc[weather['最高温'] - weather['最低温'] <= 100, '温差max'] = "温差小"
weather.value_counts()

# 四、Pandas数据统计函数
'''
1. 汇总统计
2. 唯一去重和按值计数
3. 相关系数和协方差
'''
# 4.1 汇总类统计
weather.describe()
weather['最低温'].mean()
weather['最低温'].min()
weather['最低温'].max()

# 4.2 唯一去重和按值计数
weather['级别'].unique()
weather['级别'].value_counts()

# 4.3 相关系数和协方差
# 协方差矩阵
weather.cov()
weather['最低温'].cov(weather['最高温'])
# 相关系数矩阵
weather.corr()
weather['级别'].corr(weather['最高温'] - weather['最低温'])

# 五、缺失值处理
import pandas as pd

fpath = '/Users/dohozou/Desktop/Code/gitCode/Pandas/dataFiles/'
dfna = pd.read_excel(fpath + 'Na_Null.xlsx', skiprows=2)

# 检查空值
dfna.isnull()
dfna["分数"].isnull()
dfna["分数"].notnull()
# 筛选没有空分数的所有行
dfna.loc[dfna["分数"].notnull(), :]
dfna.loc[dfna["科目"].notnull(), :]

# 删除全为空值的列
dfna.dropna(axis="columns", how="all", inplace=True)
# 删除全为空值的行
dfna.dropna(axis="index", how="all", inplace=True)

# 将分数为空的值填充为0
dfna.fillna({"分数": 0})
dfna.loc[:, "分数"] = dfna["分数"].fillna(0)

# 将姓名填充,ffill: forward fill
dfna.loc[:, "姓名"] = dfna["姓名"].fillna(method="ffill")

# 保存数据
dfna.to_excel(fpath + "notNa.xlsx", index=False)


# 六、SettingWithCopyWarning报警
import pandas as pd

fpath = '/Users/dohozou/Desktop/Code/gitCode/Pandas/dataFiles/'
weather = pd.read_excel(fpath + 'weather_xlsx.xlsx', sheet_name=0)
weather.head()
# 替换温度后面的文本
weather.loc[:, '最低温'] = weather['最低温'].str.replace('"C', '').astype('int32')
weather.loc[:, '最高温'] = weather['最高温'].str.replace('"C', '').astype('int32')
weather.head()

condition = weather['日期'].astype("str").str.startswith("2023-02")

# 6.1. 新增温差列
# weather[condition]['温差'] = weather['最高温'] - weather['最低温'] # SettingWithCopyWarning报警
# 解决方法一
weather.loc[condition, '温差'] = weather['最高温'] - weather['最低温']
# 解决方法二
df_new = weather[condition].copy()
df_new['wencha'] = df_new['最高温'] - df_new['最低温']

# 七、数据排序
weather.sort_values()












