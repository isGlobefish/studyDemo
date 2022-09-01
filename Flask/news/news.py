'''
逝者如斯夫, 不舍昼夜 -- 孔夫子
@Auhor    : Dohoo Zou
Project   : gitCode
FileName  : news.py
IDE       : PyCharm
CreateTime: 2022-09-02 02:32:33
'''

# 导入蓝图
from flask import Blueprint

"""
实例化蓝图对象
第一个参数：蓝图名称
第二个参数：导入蓝图的名称
第三个参数：蓝图前缀，该蓝图下的路由规则前缀都需要加上这个
"""
blueprint = Blueprint('news', __name__, url_prefix="/news")


# 用蓝图注册路由
@blueprint.route("/society/")
def society_news():
    return "社会新闻板块"


@blueprint.route("/tech/")
def tech_news():
    return "新闻板块"
