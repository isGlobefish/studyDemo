'''
逝者如斯夫, 不舍昼夜 -- 孔夫子
@Auhor    : Dohoo Zou
Project   : gitCode
FileName  : products.py
IDE       : PyCharm
CreateTime: 2022-09-02 02:32:44
'''

from flask import Blueprint

blueprint = Blueprint("products", __name__, url_prefix="/products", template_folder='templates', static_folder='static')


@blueprint.route("/car")
def car_products():
    return "汽车产品版块"


@blueprint.route("/baby")
def baby_products():
    return "婴儿产品版块"
