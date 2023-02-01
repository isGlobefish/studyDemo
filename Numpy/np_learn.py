'''
逝者如斯夫, 不舍昼夜 -- 孔夫子
@Auhor    : Dohoo Zou
Project   : gitCode
FileName  : np_learn.py
IDE       : PyCharm
CreateTime: 2023-01-31 20:48:17
'''
import numpy as np


# 数组与矩阵
# 二维数据和矩阵一样
# 两个数组相加
def add_array(n):
    a = [i ** 3 for i in range(1, n + 1)]
    b = [i ** 2 for i in range(1, n + 1)]
    c = []
    for i in range(n):
        c.append(a[i] + b[i])
    return c


print(add_array(4))


# 两个数组相加-np方法
def add_array(n):
    a = np.arange(1, n + 1) ** 3
    b = np.arange(1, n + 1) ** 2
    return a + b


print(add_array(4))

# -------------------------------------------
# 一、创建数组的方法
# -------------------------------------------
# 1.1 array()
a = np.array([1, 2, 3, 4, 5])
a.dtype
# 1.2
b = np.array(range(6))
# 修改数据类型
b = np.array(range(6), dtype=float)
b = np.array(range(6), dtype='float32')
'''
array的属性
shape：返回元组 几行几列 shape(2,3)：两行三列 shape(2, 2, 3): 两个两行三列
ndim: 返回一个数字 array维度数目
size：返回一个数字 array所有数据元素个数
dtype：返回array中元素的数据类型
'''
a.shape

# 1.3 arange() 返回数组 推荐使用
c = np.arange(1, 6)
c.dtype

# 1.4 ones() 创建全是1的数组
# ones_like() 形状不变 元素全改为1
np.ones(3)
np.ones((3,))
np.ones((2, 3))

# 1.5 zeros() 创建全是0的数组
# zeros_like() 形状不变 元素全改为0
np.zeros((2, 3))

# 1.6 full() 创建全是指定的字符
# full_like() 形状不变 元素全改为指定字符
np.full(3, 520)
np.full((2, 4), 520)

# 1.7 多维数组
aa = np.array([[1, 2, 3], [4, 5, 6]])
aa.ndim

np.ones((2, 3, 4))

# 1.8 使用random生成随机数组
np.random.random()  # 一个随机数[0,1)
np.random.random(3)  # 三个随机数[0,1)
np.random.random(3, 2)  # 三行两列随机数[0,1)
np.random.random(3, 2, 4)  # 3个 两行四列 一个随机数[0,1)

np.round(12.3455, 2)  # 保留几位小数

# -------------------------------------------
# 二、重塑数组
# -------------------------------------------
# 2.1 reshape()
# 把一维数组变成两维
np.arange(10).reshape(2, 5)

# 数组计算
# 形状一样的数组
a = np.arange(10).reshape(2, 5)
b = np.random.randn(2, 5)
a + b
# 形状不一样的数组
a + 1
# 行相同 则每一行运算 列相同 则每一每一列运算
# 什么维度的数组之间不能运算？
# 首尾维度一样的数组之间可以运算 除此不能

# 基础索引与切片
# 一维
a[2]
a[-1]
a[2:4]
# 二维
a[1, 2]  # 第2行第3列元素
a[-1]  # 取最后一行元素
a[:, 2]  # 取第3列全部元素
a[0:2, 2:4]

# 1.5 布尔索引
# 一维数组
a = np.arange(10)
筛选 = a > 5
a[筛选]

a[a <= 5] = 1
a[a > 5] = 0
a[a > 5] += 520

# 二维数组
b = np.arange(1, 21).reshape(4, 5)
# 返回一维数组
b[b > 10]
# 第4列元素全部改为520
b[:, 3] = 520
# 第4列元素中大于等于15的元素全部改为520
b[:, 3][b[:, 3] >= 15] = 520

# 条件组合：偶数或者小于7的数
a = np.arange(10)
条件 = (a % 2 == 0) | (a < 7)
# 条件 = (a % 2 == 0) & (a < 7)
a[条件]

# 1.6 神奇索引
a = np.arange(10)
# 第几个
a[[2, 3, 5]]

b = np.arange(36).reshape(9, 4)
# 第几行
b[[4, 5, 6, 7]]
# 第几行第几列
b[[1, 2, 3, ], [3, 2, 1]]
b[:, [3, 2, 1]]

# 拿出数组前三个最大的值
c = np.random.randint(1, 100, 10)
下标 = c.argsort()[-3:]
print(下标)
c[下标]

# 1.7 数组的轴与转置
# 0是行 1是列 2是纵深
d = np.arange(16).reshape(2, 8)
# 转置
d.transpose()

# 1.8 随机数
import random

random.seed(100)  # 随机种子
random.random()
random.random()
random.random()

# 1.8.1 rand() 0-1之间的数
np.random.rand(3)  # 一维
np.random.rand(2, 3)  # 二维
np.random.rand(2, 3, 4)  # 三维

# 1.8.2 randn() 返回标准正态分布随机数 平均数0 方差1
np.random.randn(3)  # 一维
np.random.randn(2, 3)  # 二维
np.random.randn(2, 3, 4)  # 三维

# randint() 随机整数
np.random.randint(1, 10, size=(5,))  # 一维
np.random.randint(1, 10, size=(2, 5))  # 二维
np.random.randint(1, 20, size=(2, 5, 2))  # 三维

# random 生成0.0到1.0的随机数
np.random.random(size=(5,))  # 一维
np.random.random(size=(2, 5))  # 二维
np.random.random(size=(2, 5, 2))  # 三维

# choice() 从一维数组中生成随机数
np.random.choice(5, 3)  # 0-4中随机抽三个数
np.random.choice(5, (2, 3))  # 0-4中随机抽三个数组成2*3的数组
np.random.choice([1, 2, 4, 6, 7], 3)
np.random.choice([1, 2, 4, 6, 7], (2, 3))

# shuffle() 把一个数组随机排列
# 一维数组位置随机排序
# 二维数组按行随机排序
# 三维数据按块随机排序
a = np.arange(10)
np.random.shuffle(a)

#  permutation() 和shuffle差异不大
np.random.permutation(10)  # 一维
a = np.arange(9).reshape(3, 3)
np.random.permutation(a)  # 和shuffle的差异是原来的数组没有发生变化

# normal() 生成正态分布数字
np.random.normal(1, 10, 10)  # 平均数1 方差10 10个数

# uniform() 均匀分布
np.random.uniform(1, 10, 10)  # 1到10之间 10个数
np.random.uniform(1, 10, (2, 3))  # 1到10之间 2*3个数

# 数学和统计方法
'''
sum
average               加权平均
prod                  积
mean
std var
min max
argmin argmax argsort 最大最小值下标、排序下标
cumsum                累计和
cumprod               累积积
median
precentile            0-100百分位数
quantile              0-1分位数
bincount              统计非负整数个数
'''
# 众数 np没有直接方法 可以间接获得
a = [1, 1, 2, 3, 4, 4, 4, 4, 5, 6]
counts = np.bincount(a)
np.argmax(counts)

# pandas 计算众数
'Series'.mode()

# Numpy的axis参数
# axis=0代表行 列求和  axis=1代表列 行求和


# 将条件逻辑作为数组操作
a = np.array([[1, 3, 5], [2, 3, 6], [5, 7, 9]])
a > 3
np.where(a > 3, 520, 120)  # 大于3返回520 否则返回120
(a > 3).sum()  # 大于3的个数

# any 至少一个True
# all 所有都是True
(a > 3).any()
(a > 3).all()

b = np.array([2, 1, 3, 4, 5, 0, 6, 7]).reshape(2, 4)
np.sort(b)  # 默认按最后的轴排序
np.sort(b, axis=0)  # 按行排序

# argsort() 从大到小的索引
np.argsort(a)  # 从大到小
np.argsort(-a)  # 从小到大

# lexsort() 按照键值的字典序排序

# unique() 唯一值/去重
'''
intersect1d(x,y) 交集并排序
union1d(x,y)     并集并排序
in1d(x,y)        x是否在y 返回布尔值数组
setdiff1d(x,y)   在x中但不在y中
setxord(x,y)     不是xy交集的其他部分
'''
