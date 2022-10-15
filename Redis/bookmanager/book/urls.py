'''
逝者如斯夫, 不舍昼夜 -- 孔夫子
@Auhor    : Dohoo Zou
Project   : gitCode
FileName  : urls.py
IDE       : PyCharm
CreateTime: 2022-10-14 18:01:08
'''
from django.urls import path
from book import views

urlpatterns = [
    # 第一个参数是正则 第二个参数是函数名，注意：不是函数()
    path('index/', views.index)
]
