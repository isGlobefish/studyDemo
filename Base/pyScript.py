'''
逝者如斯夫, 不舍昼夜 -- 孔夫子
@Auhor    : Dohoo Zou
Project   : gitCode
FileName  : pyScript.py
IDE       : PyCharm
CreateTime: 2022-08-31 00:32:02
'''

import logging

# a = 5
# b = 0
# try:
#     c = a / b
# except Exception as e:
# 下面三种方式三选一，推荐使用第一种
# logging.exception("Exception occurred")
# logging.error("Exception occurred", exc_info=True)
# logging.log(level=logging.ERROR, msg="Exception occurred", exc_info=True)

# selectNum = (lambda inp: int(inp) if inp.isdigit() else logging.exception('都是垃圾'))(input('「」_'))


# from Log.log import loggers
#
# loggers.debug('debug')
# loggers.critical('1111111111')

import os
from pathlib import Path

print(os.getcwd())
print(Path.cwd())

print(os.path.dirname(os.path.dirname(os.getcwd())))
print(Path.cwd().parent.parent)

print(os.path.join(os.getcwd(), 'Config', 'requirements.txt'))
print(Path.cwd().joinpath(*['Config', 'requirements.txt']))

from loguru import logger

logger.debug('123')
# coding:utf-8
from loguru import logger

logger.add("../Log/loguru_{time}.log", rotation="500MB", encoding="utf-8", enqueue=True, compression="zip", retention="10 days")
logger.info("中文test")
logger.debug("中文test")
logger.error("中文test")
logger.warning("中文test")

from Util.operate_yaml import YamlUtil

YamlUtil().read_yaml('/Users/dohozou/Desktop/Code/gitCode/Config/mysql.yaml')

import os

print(os.listdir(os.getcwd()))

print(os.path.realpath(__file__))

print(os.path.abspath('.'))

print(os.path.abspath(os.path.realpath(__file__)))

print(os.path.split(os.path.realpath(__file__)))

print(os.path.join(os.getcwd(), 'Config', 'mysql.yaml'))

print(os.system('pip3 list'))

from tqdm import tqdm

for i in tqdm(range(10000000)):
    temp = ['你好'] * 2000


