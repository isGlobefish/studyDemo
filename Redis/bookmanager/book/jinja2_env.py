'''
逝者如斯夫, 不舍昼夜 -- 孔夫子
@Auhor    : Dohoo Zou
Project   : gitCode
FileName  : jinja2_env.py
IDE       : PyCharm
CreateTime: 2022-10-20 22:06:03
'''

from django.template.defaultfilters import date
from jinja2 import Environment


def environment(**option):
    env = Environment(**option)

    env.globals.update({
        'data': date,
    })
    return env