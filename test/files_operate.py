'''
逝者如斯夫, 不舍昼夜 -- 孔夫子
@Auhor    : Dohoo Zou
Project   : gitCode
FileName  : files_operate.py
IDE       : PyCharm
CreateTime: 2022-08-28 22:02:57
'''
import os

with open('../test/script.txt', 'a+', encoding='utf-8', errors='ignore') as f:
    # f.write('我爱世界, 更爱中国')
    # f.writelines('新的一行')
    print(f.read())
    f.close()

path = os.getcwd()

for dirpath, dirnames, filenames in os.walk(path):
    print(dirpath)
    print(dirnames)
    print(filenames)


print(os.sep)