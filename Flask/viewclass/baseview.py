'''
逝者如斯夫, 不舍昼夜 -- 孔夫子
@Auhor    : Dohoo Zou
Project   : gitCode
FileName  : baseview.py
IDE       : PyCharm
CreateTime: 2022-09-02 14:52:57
'''

from flask import views, render_template


class BaseView(views.View):
    # 如果子类忘记定义 get_template 就会报错
    def get_template(self):
        raise NotImplementedError()

    # 如果子类忘记定义 get_data 就会报错
    def get_data(self):
        raise NotImplementedError()

    def dispatch_request(self):
        template = self.get_template()
        data = self.get_data()
        return render_template(template, **data)
