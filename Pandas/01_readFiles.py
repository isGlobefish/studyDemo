'''
逝者如斯夫, 不舍昼夜 -- 孔夫子
@Auhor    : Dohoo Zou
Project   : gitCode
FileName  : 01_readFiles.py
IDE       : PyCharm
CreateTime: 2022-11-26 18:11:26
'''
# -------------------------------------------
# 零、读取数据
# -------------------------------------------
import pandas as pd

fpath = '/Users/dohozou/Desktop/Code/gitCode/Pandas/dataFiles/'

'''
读取数据的三种方法
pd.read_csv() csv txt
pd.read_excel() excel
pd.read_mysql()

pd.read_table()
'''

# 读取csv文件数据
'''
# header=0: 表头是第一行
# header=None：没有表头
# names=[]：自定义表头
# index_col=''/[]: 指定列作为索引, 多个索引列用列表
# skiprows=[]:跳过指定行
# nrows=10：读取多少行数据
# encoding='utf-8'

to_csv()
'''
# data_csv = pd.read_csv(fpath + 'pandas_csv.csv', header=None, sep='\t',names=['A','B','C'], index_col='xxx')
data_csv = pd.read_csv(fpath + 'pandas_csv.csv', header=0)
'''
查看数据各种属性
head tail shape columns index dtypes
'''
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
'''
header=0/None: 表头
names=[]: 自定义列名
index_col=''/[] # 指定索引列
inplace=True # 默认False True表示在原数据上修改
usecols='F:I' # 指定使用范围列的值
dtype={'序号':str, '性别':str, '日期':str} # 规定每列的数据类型
parse_dates=['出生日期'] # 把该列数据改为日期型

to_excel()
'''
data_xlsx = pd.read_excel(fpath + 'pandas_xlsx.xlsx', sheet_name=0)

# 既有\t 也有\n之类的，sep使用正则
pd.read_table(fpath + 'pandas_csv.csv', sep='\s+')

# 读取数据库数据
# 方法一：pymysql读取数据库中的数据
import pymysql

conn = pymysql.connect(
    host='localhost',
    port=3306,
    user='root',
    password='88888888',
    database='my_test',
    charset='utf8mb4'
)
# 创建游标对象
cursor = conn.cursor()

getData = cursor.execute('select * from day20')
conn.commit()
'''
返回值是元组tuple
fetchone fetchmany fetchall
'''
allData = cursor.fetchall()
# 关闭游标
cursor.close()
# 关闭连接
conn.close()

# -------------------------------
# pymsql + pandas
pd.read_sql('查询数据语句', con=conn)

# 方法二：sqlalchemy读取数据库中的数据
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

# 获取所有的设置值
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

# -------------------------------------------
# 一、数据结构
# -------------------------------------------
'''
Dataframe：二维数据 整个表格
Series: 一维数据 一行或者一列
'''
# 1. Pandas的数据结构Series
import pandas as pd
import numpy as np

s1 = pd.Series([1, 'a', 3.14, 7])
# 获取Series的索引
s1.index
# 获取Series的值
s1.values

# 创建具有标签索引的Series
s2 = pd.Series([1, 'a', 3.14, 7], index=['a', 'b', 'c', 'd'])
s2.index
s2.values
list1 = [1, 'a', 3.14, 7]
list2 = ['a', 'b', 'c', 'd']
s2 = pd.Series(list1, index=list2)
s2.sort_values()
s2.isnull()
s2.notnull()

# 使用字典创建Series
dict = {'a': 100, 'b': 200, 'c': 300}
s3 = pd.Series(dict)
# 查询与字典操作类似
print(s3['a'])
print(s3[['a', 'b']])
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

# 保存数据，去掉/指定索引列
df.reset_index('one')
df.to_excel(fpath, index=None)

'''
1、如果查询一行或者一列 返回的数据类型是Series
2、如果查询多行或者多列 返回的数据类型是Dataframe
查询列或者多列使用key/列名 查询行或者多行使用loc
'''
# 查询一列
df['one']
# 查询多列
df[['one', 'two']]
type(df['one'])  # Series
type(df[['one', 'two']])  # dataframe
# 查询一行，结果是一个pd.Series
df.loc[1]
# 查询多行
df.loc[0:3]

# -------------------------------------------
df1 = pd.DataFrame([[1, 2, 3], [4, 5, 6], [7, 8, 9]], columns=['a', 'b', 'c'])
df1[1][1]
df1.loc[1, 'b']
# 单元格从0开始,与列明无关
df1.iloc[1][1]

# 多个series数据生成dataframe
s1 = pd.Series(['成龙', '李连杰', '林青霞'], index=[1, 2, 3], name='姓名')
s2 = pd.Series(['男', '男', '女'], index=[1, 2, 3], name='性别')
s3 = pd.Series(['60', '62', '59'], index=[1, 2, 3], name='年龄')
# 方法一
df = pd.DataFrame({s1.name: s1, s2.name: s2, s3.name: s3})
# 方法二
df = pd.DataFrame([s1, s2, s3])

# dataframe常用方法
'''
index
columns
dtypes
head()
tail()
shape()
fillna() # 填充数值
replace()
isnull()
notnull()
unique()
reset_index(drop=False)
sort_index()
sort_values()
'''

# -------------------------------------------
# 二、数据查询
# -------------------------------------------
'''
数据查询的五种方法：数值 列表 区间 条件 函数

1. df.loc方法，根据行、列标签查询
2. df.iloc方法，根据行、列的数字位置查询
3. df.where方法
4. df.query方法

.loc既能查询，又能覆盖写入 重点！！！
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

# 2.1 使用「单个label值」查询数据
weather.loc['2023-01-01', "最低温"]  # 得到单个值
weather.loc['2023-01-01', ["最低温", "最高温"]]  # 得到单个series
weather.loc[['2023-01-01', '2023-01-02', '2023-01-03'], "最低温"]  # 得到series
weather.loc[['2023-01-01', '2023-01-02', '2023-01-03'], ["最低温", "最高温"]]  # 得到dataframe

# 2.2 使用「数值区间」进行范围查询
weather.loc['2023-01-01':'2023-01-05', '最低温']
# 列index按区间
weather.loc['2023-01-01', '最低温':'级别']
# 行和列都按区间查询
weather.loc['2023-01-01':'2023-01-05', '最低温':'级别']

# 2.3 使用「条件表达式」查询
weather.loc[weather['最低温'] < -30, :]
# 复杂查询使用&
weather.loc[(weather['最低温'] < -30) & (weather['最高温'] > 50) & (weather['级别'] == 1)]

# 2.4 使用「函数」查询
weather.loc[lambda df: (df['最低温'] < -30) & (df['最高温'] > 50), :]


def query_weather(df):
    return df.index.astype('str').str.startswith('2023-02')


weather.loc[query_weather, :]

# -------------------------------------------
# 二、数据筛选（新）
# -------------------------------------------
# 2.1 筛选范围行的数据
weather.loc[1:3]
# 2.2 只选男性这一类的数据
条件 = weather['性别'] == '男'
weather[条件]
# query()
条件 = "性别 == '男' and 总分 >= 150"
weather.query(条件)
条件 = "性别 == '女' and 60 <= 总分 <= 150"
weather.query(条件)
# 2.3 文本以xx开头结尾的数据
weather['姓名'].str.startswith('王')
weather['姓名'].str.endswith('王')
# 2.4 包含某字符的数据
weather['地址'].str.contains('北京市')
# 正则
条件 = weather['地址'].str.contains('[a-cA-C]座')
weather[条件]
# 2.5 index_col = '出生日期' parse_dates=['出生日期']
weather['1989']
weather['1989-10']
# 使用truncate()前要排序数据
df = weather['出生日期'].sort_values()  # df = weather.sort_values('出生日期')
df.truncate(before='1990-01-01')  # 在该日期前
df.truncate(after='1990-01-01')  # 在该日期后
# 该日期范围内的数据
weather['1990':'2000']
weather['1990-01-02':'2000-05-06']

# 注意：读取文件时，不能把 出生日期 该列设为index_col
条件 = (
    '@weather.出生日期.dt.year > 1980 and'
    '@weather.出生日期.dt.year < 1990'
    'and 性别 == "男"'
)
weather.query(条件)

# -------------------------------------------
# 三、新增数据列
# -------------------------------------------
'''
1. 直接赋值
2. df.apply方法
3. df.assign方法
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

# -------------------------------------------
# 三、新增数据列（新）
# -------------------------------------------
# 查看风向这列的字数，是len，apply传入函数名即可，并不是len()
weather['风向'] = weather['风向'].apply(len)

import pandas as pd
import numpy as np

df = pd.DataFrame(np.arange(9).reshape(3, 3), columns=list('xyz'), index=list('abc'))

# 每个元素都平方根
df.apply(np.square)
# 条件平方根
df.apply(lambda m: np.square(m) if m.name == 'x' else m)  # 列
df.apply(lambda m: np.square(m) if m.name == 'a' else m, axis=1)  # 行

# -------------------------------------------
# 三、删除数据列（新）
# -------------------------------------------
df.drop(1)  # 删除第一行，默认aixs=0
df.drop(labels=[1, 3])  # 删除1-2行，默认aixs=0

df.drop(labels=['语文', '数学'], axis=1, inplace=True)  # 删除两列 axis=1
df.drop(labels=['A', 'F'], axis=1, inplace=True)  # 删除两列 axis=1

# -------------------------------------------
# 四、Pandas数据统计函数
# -------------------------------------------
'''
1. 汇总统计
2. 唯一去重和按值计数
3. 相关系数和协方差
'''
# 4.1 汇总类统计
weather.describe()  # 统计描述
weather['最低温'].mean()
weather['最低温'].min()
weather['最低温'].max()
weather['最低温'].median()  # 中位数
weather['最低温'].count()  # 非空值个数
weather['最低温'].std()  # 标准差
weather['最低温'].var()  # 方差
weather['最低温'].mad()  # 平均绝对方差
weather['最低温'].mode()  # 众数
weather['最低温'].idxmin()  # 最小值行索引
weather['最低温'].idxmax()  # 最大值行索引

# 4.2 相关系数和协方差
# 协方差矩阵
weather.cov()
weather['最低温'].cov(weather['最高温'])
# 相关系数矩阵
weather.corr()
weather['级别'].corr(weather['最高温'] - weather['最低温'])

# -------------------------------------------
# 四、删除重复值（新）
# -------------------------------------------
# 4.1 唯一去重和按值计数
weather['级别'].unique()
weather['级别'].value_counts()

'''
Dataframe.drop_duplicates(subset=None, keep='first', inplace=False)
参数
subset：用来指定特定的列，默认是所有列
keep：指定保存重复值的方法
    first：保留第一次出现的值
    last：保留最后出现的值
    False：删除所有重复值 留下没有出现过的重复的
inplace：是直接在原来数据上修改还是保留一个副本
'''
weather.drop_duplicates(subset=['级别'], keep='first')  # 去重
weather.duplicated(subset=['级别'], keep='first')  # 保留重复项

# -------------------------------------------
# 四、算数运算与数据对齐（新）
# -------------------------------------------
# 两列中存在Nan值时 如何做运算
df['新列'] = df['店1'].fillna(0) + df['店2'].fillna(0)  # 方法一
'''
add      radd      加法
sub      rsub      减法
div      rdiv      除法
floordiv rfloordiv 整除
mul      rmul      乘法
'''
df['新列'] = df['店1'].add(df['店2'], fill_value=0)  # 方法二

# 出现无穷大inf -inf时 如何解决 把所有的无穷大转换为空
pd.options.mode.use_inf_as_na = True

# -------------------------------------------
# 五、缺失值处理
# -------------------------------------------
'''
空值是缺失值
缺失值不仅仅是空值
'''
import pandas as pd

fpath = '/Users/dohozou/Desktop/Code/gitCode/Pandas/dataFiles/'
dfna = pd.read_excel(fpath + 'Na_Null.xlsx', skiprows=2, index=False)

# 5.1 检查空值
dfna.isnull()
dfna["分数"].isnull()
dfna["分数"].notnull()
# 5.2 筛选没有空分数的所有行
dfna.loc[dfna["分数"].notnull(), :]
dfna.loc[dfna["科目"].notnull(), :]

'''
axis=0：删除包含缺失值的行
axis=1：删除包含缺失值的列

how='any'：只要有缺失值出现 就删除该行或列
how='all'：所有值都缺失 就删除该行或列

thresh：axis中至少有thresh个非缺失值 否则删除

subset：list 规定删除的列范围
'''
# 5.3.1 删除全为空值的列
dfna.dropna(axis="columns", how="all", inplace=True)
# 5.3.2 删除全为空值的行
dfna.dropna(axis="index", how="all", inplace=True)

# -------------------------------------------
# 五（新）、自动填充
# -------------------------------------------
# 将分数为空的值填充为0
dfna.fillna({"分数": 0})
dfna.loc[:, "分数"] = dfna["分数"].fillna(0)

# 将姓名填充,ffill: forward fill
'''
method：{'backfill', 'bfill', 'pad', 'ffill', None} 默认None
pad/fill: 前一个数值填充
backfill/bfill: 后一个数值填充

limit： 限定最多填充多少个

左右填充只需把前后填充axis改变就是
'''
dfna.loc[:, "姓名"] = dfna["姓名"].fillna(method="ffill")

'''
in not in at的使用
'zou' in ['zou', 'de', 'hao']
'zip' not in ['zou', 'de', 'hao']
series.at[i] = i + 1
'''

# 保存数据
dfna.set_index('序号', inplace=True)
dfna.to_excel(fpath + "notNa.xlsx", index=False)

# -------------------------------------------
# 六、SettingWithCopyWarning报警
# -------------------------------------------
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

# -------------------------------------------
# 七、数据排序默认升序Ture
# -------------------------------------------
weather['最低温'].sort_values(ascending=False)  # Series
weather.sort_values(by='最低温', ascending=False)  # Dataframe
# 多列排序，分别指定排列方式
weather.sort_values(by=['最低温', '风向'], ascending=[False, True], inplace=True)  # Dataframe

# 按索引排序
# by=1对第二行排序 axis=1表示对列变
weather.sort_index(by=1, inplace=True, ascending=False, axis=1)

# -------------------------------------------
# 八、字符串处理
# -------------------------------------------
'''
在series属性上调用函数
只能在字符串列上使用，不能在数字列上使用
Dataframe上没有str属性和处理方法
Series.str并不是python的原生字符串，而是自己的一套方法，不过大部分和原生str相似
'''
# 8.1 astype() int32 int64 float32 float64 str
weather['级别'].astype('str').str.isnumeric()
weather['级别'].astype('str').str.len()
weather['级别'].astype('str').str.startswith('2023-02')
weather['日期'].astype('str').str.replace('-', '')[0:]
weather['日期'].str.replace('-', '').str.slice(0, 6)
# 8.2 cat() 分割
weather['级别'].astype('str').str.cat()  # 连成一串
weather['级别'].astype('str').str.cat(sep=',')  # 以,分割
weather['级别'].astype('str').str.cat(['-'] * len(weather['级别']))  # 以-分割
weather['级别'].astype('str').str.cat(['-'] * len(weather['级别']), sep='^', na_rep='没有')  # Nan替换
# 8.3 split() 分割
weather['级别'].str.split('血', n=2, expand=True)
# 8.4 partition() 只分割第一个
'BbBbBBBBBB'.str.partition('b')  # 从左到右
'BbBbBBBBBB'.str.rpartition('b')  # 从右到左
# 8.5 get() 获取指定位置的字符
weather['级别'].str.get(2)
# 8.6 slice(m, n) 获取指定范围的字符
weather['级别'].str.slice(0, 3, 2)
# 8.7 slice_replace() 筛选后的替换
weather['级别'].str.slice_replice(0, 3, '新字符')
# 8.8 join() 连接字符
weather['级别'].str.join('新字符')
# 8.9 contains() 字符串是否包含指定字符 指定na为False或者其他字符串
weather['级别'].str.contains('ok', na='False')
# 8.10 startswith() 以什么开头
weather['级别'].str.startswith('a')
# 8.11 staswith() 以什么开头
weather['级别'].str.endswith('z')
# 8.12 repeat() 重复多少次
weather['级别'].str.repeat(3)
# 8.13 pad() 补齐多少个字符 side={'left', 'right', 'both'}
'''
center()等价于side='both'
ljust()等价于side='left'
rjust()等价于side='right'
'''
weather['级别'].str.pad(5, fillchar='*', side='both')
# 8.14 zfill() 填充0
'123'.zfill(5)
# 8.15 编码encode() 解码decode()
weather['级别'].str.encode('utf-8').str.decode()
# 8.16 strip() 去掉指定字符串 lstrip()从左 rstrip()从右
weather['级别'].str.strip('abcd')
# 8.17 get_dummies() 矩阵形式呈现 1代表有 0代表无
weather['级别'].str.get_dummies('距')
# 8.18 translate() 按照指定部分替换
字典 = str.maketrans({'大': 'da', '小': 'xiao'})
weather['级别'].str.translate(字典)
# 8.19 find() 查找字符第一次出现的位置 第二个参数是查找起始位置 找不到返回-1
# rfind() 从右查 index() 找不到报错 rindex() 从右边找 找不到报错
weather['日期'].astype('str').str.find('-', 5)

'''
其他字符串处理
lower() # 所有字符变成小写
upper() # 所有字符变成大写
title() # 每个单词的首字母大写
capitalize() # 第一个字母大写
swapcase() # 大小写交换

判断 返回True或False
isalpha() # 是否全是字母
isnumeric() # 是否全是数字
isalnum() # 是否全是字母或者数字
# isdecima只能用于Unicode数字
# isdigt只能用于Unicode数字，罗马数字
# isnumeric只能用于Unicode数字，罗马数字,汉字数字
# isnumeric用途较广泛，普通阿拉布数字三者无异

isspace() # 是否全是空格
islower() # 是否全是小写
istitle() # 是否每个单词首字母大写其他小写
'''

# 8.20 正则表达式
weather['日期'].str.replace('[年月日]', '')
weather['日期'].str.extract(r'(\d{4}\d{2}\d{2})', expand=False)
weather['日期'].str.match('.{2}激', na=False)
# 分组捕获 ()表示需要分组
weather['日期'].str.extract('(\d{4})-(\d{2})-(\d{2})')
# 把2022-01-03变成03/02/2022
weather['日期'].str.extract('(\d+)-(\d+)-(\d+)', r'\3/\2/\1')

# -------------------------------------------
# 九、axis参数，指定那个参数，那个要动起来原则
# -------------------------------------------
'''
0: index cross rows
1: columns cross columns
'''
import pandas as pd
import numpy as np

df = pd.DataFrame(
    np.arange(12).reshape(3, 4),
    columns=["A", "B", "C", "D"]
)

# 9.1 单列drop, 删除列
df.drop("A", axis=1)

# 删除行
df.drop(1, axis=0)

df.mean(axis=0)
df.mean(axis=1)

# -------------------------------------------
# 十、索引index的用途
# -------------------------------------------
df.count()
df.set_index("A", inplace=True, drop=False)
# 查询A=4的数据
df.loc[df["A"] == 4, :]
df.loc[4]  # 通过索引搜索性能快

'''
索引是否递增
Datafrme.index.is_monotonic_increasing
索引是否唯一
Datafrme.index.is_unique

魔法函数-计时
%timeit dataframe.loc[index]
'''

# index数据自动对齐
s1 = pd.Series([1, 2, 3], index=list("abc"))
s2 = pd.Series([1, 2, 3], index=list("bcd"))
s1 + s2

'''
Categoricallindex，基于分类数据的index，提升性能等
MultiIndex，对维索引，用于groupby对维聚合后结果等
DatetimeIndex，时间类型索引，强大的日期和时间的方法支持；
'''

# -------------------------------------------
# 十一、Merge语法
# -------------------------------------------
'''
连接查询
inner join
left join
right join
outer join

concat：可以沿一条轴将多个对象连接到一起
merget：可以根据一个或多个键将不同的dataframe中的行连接起来
join：inner 交集 outet是并集
'''
import pandas as pd
import numpy as np

df1 = pd.DataFrame({'姓名': ['A', 'B', 'C', 'D', 'E', 'F'], '手速': np.arange(6)})
df2 = pd.DataFrame({'姓名': ['A', 'B', 'E'], '脚速': [1, 2, 3]})
# how：inner left right outer，默认inner连接
# on：可以多个key连接
# suffixes：后缀参数suffies=['_x', '_y']
df.merge(df1, df2, on='姓名', how='outer')

# 通过索引连接
df1 = pd.DataFrame({'姓名': ['A', 'B', 'C', 'D', 'E', 'F'], '手速': np.arange(6)})
df2 = pd.DataFrame({'数据': ['A', 'B', 'E']}, index=['a', 'b', 'c'])
pd.merge(df1, df2, left_on='姓名', right_on=True, how='outer')

# -------------------------------------------
# 十一（新）、join
# -------------------------------------------
# 组合多个dataframe数据，inner是交集 outer是并集
df1.join(df, df2)

# -------------------------------------------
# 十二、Concat语法
# -------------------------------------------
# ignore_index=True忽略索引 会重新生成索引
pd.concat(['dataframe', 'dataframe'], ignore_index=True, axis=0)
pd.concat(['dataframe', 'series', 'series'], ignore_index=True, axis=1)

# axis=0行变 axis=1列变
arr = np.arange(9).reshape(3, 3)
arr1 = np.concatenate([arr, arr], axis=1)
arr2 = np.concatenate([arr, arr], axis=0)

s1 = pd.Series([0, 1, 2], index=['A', 'B', 'C'])
s2 = pd.Series([3, 4], index=['D', 'E'])
# 默认行变 排列方式默认index索引
c = pd.concat([s1, s2], axis=0, sort=True)
# 可以添加一列索引keys, 几张变添加几个
c = pd.concat([s1, s2], axis=0, sort=True, keys=['x', 'y'])
# join参数 join_axes
c = pd.concat([s1, s2], join='outer', join_axes=[s1.index])

# -------------------------------------------
# 十二（新）、append语法
# -------------------------------------------
# 默认axis=0, 可以在dataframe数据后面追加一行series数据
df1.append(s1, ignore_index=True)

# 忽略报警
import warnings

warnings.filterwarnings('ignore')

# append语法
# dataframe.append(dataframe, ignore_index=True)

# 低性能版本
for i in range(5):
    df = df.append({'A': i}, ignore_index=True)
# 高性能版本
pd.concat(
    [pd.DataFrame([i], columns=['A']) for i in range(5)],
    ignore_index=True
)

# -------------------------------------------
# 十三、excel的拆分与合并
# -------------------------------------------
import os

if not os.path.exists(fpath):
    os.mkdir(fpath)

# 数据结构dataframe.shape[0]
# dataframe.iloc[begin:end]

import os

excel_names = []
for excel_name in os.listdir(fpath):
    excel_names.append(excel_name)

# -------------------------------------------
# 十四、分组统计groupby
# -------------------------------------------
import pandas as pd
import numpy as np

df = pd.DataFrame({
    'A': ['foo', 'bar', 'foo', 'bar', 'foo', 'bar', 'foo', 'foo'],
    'B': ['one', 'one', 'two', 'three', 'two', 'two', 'one', 'three'],
    'C': np.random.randn(8),
    'D': np.random.randn(8)})

# 统计所有数据列
df.groupby('A').sum()
df.groupby(['A', 'B']).mean()
df.groupby(['A', 'B'], as_index=False).mean()  # 把AB索引列变成普通列

# 同时查看多种数据统计
df.groupby("A").agg([np.sum, np.mean, np.std])
# 查看单列结果
df.groupby("A")['C'].agg([np.sum, np.mean, np.std])
df.groupby("A").agg([np.sum, np.mean, np.std])['C']
# 不同列使用不同的聚合
df.groupby("A").agg({'C': np.sum, 'D': np.mean})

# 获取分组get_group()

# -------------------------------------------
# 十五、多层索引及其计算
# -------------------------------------------
# 分层索引MultiIndex
# unstack()把二级索引变成列ser.unstack(),ser.reset_index()
# stack()把列变成二级索引
# stock.loc[(slice(None), ['2019-10-02', '2019-10-03']), :]
import pandas as pd

dict = {'班级': ['1班', '1班', '1班', '2班', '2班', '2班', '3班', '3班', '3班'],
        '学号': ['a', 'b', 'c', 'a', 'b', 'c', 'a', 'b', 'c'],
        '分数': [1, 2, 3, 11, 22, 33, 111, 222, 333]}

df = pd.DataFrame(dict)
# 设置多层索引 或者 index_col=[0, 1]
df = df.set_index(['班级', '学号'])
# 查询1班数据
df.loc[('1班', slice(None)), :]
df.loc[(('1班', 'a'), slice(None)), :]

# 查看是否是无序型数据 如是 则需排序
df.index.is_lexsorted()  # 丢弃
df.index.is_monotonic_increasing  # 新版本
df.set_index(level='科目')

df.index.levels[0]  # 外层索引
df.index.levels[1]  # 内层索引

# 15.1 多层索引的创建
import pandas as pd
import numpy as np

# from_arrays 数组
dff = pd.MultiIndex.from_arrays([['a', 'a', 'b', 'b'], [1, 2, 1, 2]], names=['x', 'y'])
# from_tuples 元组
dff = pd.MultiIndex.from_tuples([('a', 1), ('a', 2), ('b', 1), ('b', 2)], names=['x', 'y'])
# from_product 笛卡尔积
ddf = pd.MultiIndex.from_product([['a', 'b'], [1, 2]], names=list('xy'))

index1 = pd.MultiIndex.from_product([[2012, 2022], [5, 7]], names=['年', '月'])
columns1 = pd.MultiIndex.from_product([['雪梨', '香蕉'], ['花菜', '菠菜']], names=['水果', '蔬菜'])
dfdf = pd.DataFrame(np.random.random(size=(4, 4)), index=index1, columns=columns1)

# 增加总计列
总计 = dfdf['雪梨'] + dfdf['香蕉']
总计.columns
总计.columns = pd.MultiIndex.from_product([['总计'], 总计.columns])
result = pd.concat([dfdf, 总计], axis=1)

# -------------------------------------------
# 十六、数据转换函数map、apply、applymap
# -------------------------------------------
'''
map：只能用于Series
apply：Series、Dataframe
applymap：只能Dataframe
'''
# 实例：将股票代码英文转换成中文
stock = pd.DataFrame({
    '日期': ['2023-01-01', '2023-01-02', '2023-01-03', '2023-01-04'],
    '公司': ['ALI', 'BD', 'JD', 'TX'],
})

dict_company_names = {
    "bd": "百度",
    "ali": "阿里",
    "jd": "京东",
    "tx": "腾讯"
}

# 方法1：Series.map(dict)
stock['公司中文1'] = stock['公司'].str.lower().map(dict_company_names)
# 方法2：Series.map(function)
stock['公司中文2'] = stock['公司'].str.lower().map(lambda x: dict_company_names[x.lower()])

stock['公司中文3'] = stock['公司'].str.lower().apply(lambda x: dict_company_names[x.lower()])
stock['公司中文4'] = stock.apply(lambda x: dict_company_names[x['公司'].lower()], axis=1)

# Dataframe.applymap()
stock.loc[:, ['收盘', '开盘', '高']] = stock.applymap(lambda x: int(x))

# -------------------------------------------
# 十七、每个分组应用apply函数
# -------------------------------------------
df.groupby().apply(lambda x: int(x), topn=2).head()

# -------------------------------------------
# 十八、stack和pivot实现数据透视
# -------------------------------------------
stock.pivot()

# -------------------------------------------
# 十九、apply同时添加多列
# -------------------------------------------
import pandas as pd

fpath = '/Users/dohozou/Desktop/Code/gitCode/Pandas/dataFiles/'
weather = pd.read_excel(fpath + 'weather_xlsx.xlsx', sheet_name=0)
weather.head()
# 替换温度后面的文本
weather['最低温'] = weather['最低温'].map(lambda x: int(str(x).replace('"C', "")))
weather['最高温'] = weather['最高温'].map(lambda x: int(str(x).replace('"C', "")))
weather.head()


# 同时添加多列
def my_func(row):
    return row["最高温"] - row["最低温"], (row["最高温"] - row["最低温"]) / 2


weather[["wencha", "avg"]] = weather.apply(my_func, axis=1, result_type="expand")

# -------------------------------------------
# 二十、数据替换
# -------------------------------------------
df.replace('A', '优秀', inplace=True)
# 一列多替换
# 列表 一对多或多对一 推荐使用
df['A列'].replace(['A', 'B'], ['优秀', '良好'], inplace=True)
# 字典
df['A列'].replace({'A': '优秀', 'B': '良好'}, inplace=True)
df['城市'] = df['城市'].str.replace('城八', '市')

# 正则 需添加regex=True
df.replace('[A-Z]', 88, regex=True, inplace=True)

# -------------------------------------------
# 二十一、离散化和分箱
# -------------------------------------------
import pandas as pd

year = [1992, 1983, 1992, 1932, 1973]
box = [1900, 1950, 2000]
box_name = ['50年代前', '50年代后']
# result = pd.cut(year, box, labels=False)
result = pd.cut(year, box, labels=box_name)
pd.value_counts(result)

# 等平分箱
year = [1992, 1983, 1922, 1932, 1973, 1999, 1993, 1995]
result = pd.qcut(year, q=4)
