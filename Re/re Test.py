'''
逝者如斯夫, 不舍昼夜 -- 孔夫子
@Auhor    : Dohoo Zou
Project   : gitCode
FileName  : re Test.py
IDE       : PyCharm
CreateTime: 2022-10-20 23:29:20
'''
import re

# 1. 固定字符串
text = '麦叔的身高：178，体重：168，学号：123456，密码：9527'
print(re.findall(r'123456', text))

print('-----------------------------------------')

# 2. 某一类字符
# \d 数字
print(re.findall(r'\d', text))
# \D 不是数字
print(re.findall(r'\D', text))
# \w 所有字符除标点符号
print(re.findall(r'\w', text))
# 1-5的数字
print(re.findall(r'[1-5]', text))
# []中括号的内容
print(re.findall(r'[高重号]', text))

print('-----------------------------------------')

# 3. 重复的某一类字符
# 连续多个数字的
print(re.findall(r'\d+', text))
# 0个或者1个数字
print(re.findall(r'\d?', text))
# 0个或者多个数字
print(re.findall(r'\d*', text))
# 匹配3个数字的
print(re.findall(r'\d{3}', text))
# 匹配2-4个数字的
print(re.findall(r'\d{2,4}', text))
# 匹配大于1个数字的
print(re.findall(r'\d{1,}', text))
# 匹配0-8个数字的
print(re.findall(r'\d{,8}', text))

print('-----------------------------------------')

# 4. 组合
text1 = '麦叔的手机号18812345678，他还有一个电话号码18887654321，他的爱好数字是01234567891，他的座机是：0571-52152166'
print(re.findall(r'\d{4}-\d{8}', text1))

print('-----------------------------------------')

# 5. 多种情况
print(re.findall(r'\d{4}-\d{8}|1\d{10}', text1))

print('-----------------------------------------')

# 6. 限定位置 - 开头
text2 = '18812345678，他还有一个电话号码18887654321，他的爱好数字是01234567891，他的座机是：0571-52152166'
print(re.findall(r'^1\d{10}|^\d{4}-\d{8}', text2))

print('-----------------------------------------')

# 7. 内部约束
text3 = 'baibai carcar duodui'
print(re.findall(r'(\w{3})(\1)', text3))

print('-----------------------------------------')

# 语法
text4 = '随机数01234567890，座机1：0957-1111111111。座机2：0100-0000000-9999999999'
print(re.findall(r'\d{10}|\d{4}-\d{10}|\d{4}-\d{7}\d{10}', text4))

# 忽略大小写
text5 = 'abc ABC abC'
print(re.findall(r'abc', text5, flags=re.I))

# search
m = re.search(r'abc', text5)
print(m)
# group(1) group(2) groups()全部
print(m.group())








