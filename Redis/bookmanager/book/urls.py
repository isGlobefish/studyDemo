'''
逝者如斯夫, 不舍昼夜 -- 孔夫子
@Auhor    : Dohoo Zou
Project   : gitCode
FileName  : urls.py
IDE       : PyCharm
CreateTime: 2022-10-14 18:01:08
'''
from django.urls import path, re_path
from book import views

# 总urls里面添加namespace需要到子urls声明app_name
app_name = 'book'

urlpatterns = [
    # 第一个参数是正则 第二个参数是函数名，注意：不是函数()
    path('index/', views.index, name='index'),
    path('reindex/', views.reindex, name='reindex'),
    # 位置参数
    # re_path('(\d+)/(\d+)/', views.specil_url),
    # 假如后面传参两个数字用反了怎么办？推荐使用关键字参数
    re_path('(?P<category_id>\d+)/(?P<book_id>\d+)/', views.specil_url),
    # 查询字符串
    path('userpwd/', views.query_string),
    # (body) post
    path('body_post/', views.body_post),
    # set_cookie
    path('set_cookie/', views.set_cookie),
    # get_cookie
    path('get_cookie/', views.get_cookie),
    # set_session
    path('set_session/', views.set_session),
    # get_session
    path('get_session/', views.get_session),

]
