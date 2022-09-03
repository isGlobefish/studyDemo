#!usr/bin/env python
# -*- coding:utf-8 _*-
"""
# author: 小菠萝测试笔记
# blog:  https://www.cnblogs.com/poloyy/
# time: 2021/7/11 1:47 下午
# file: 5_request_form.py
"""

from flask import Flask, request

app = Flask(__name__)


@app.route('/addUser', methods=['POST'])
def check_login():
    return {"name": request.form['name'], "age": request.form['age']}


@app.route('/addUser2', methods=['POST'])
def check_login2():
    print('form =', request.form)
    print('args =', request.args)
    return "good"


@app.route('/addUser3', methods=['POST'])
def check_login3():
    print('form =', request.form)
    print('json =', request.json)
    return "good"


@app.route('/addUser4', methods=['POST'])
def check_login4():
    return {"name": request.values['name'], "age": request.values['age']}


from flask import request, views
from functools import wraps


def check_login(original_function):
    @wraps(original_function)
    def decorated_function(*args, **kwargs):
        user = request.args.get("user")
        if user and user == 'zhangsan':
            return original_function(*args, **kwargs)
        else:
            return '请先登录'

    return decorated_function


class Page1(views.View):
    decorators = [check_login]

    def dispatch_request(self):
        return 'Page1'


class Page2(views.View):
    decorators = [check_login]

    def dispatch_request(self):
        return 'Page2'


def check_login(original_function):
    @wraps(original_function)
    def decorated_function(*args, **kwargs):
        user = request.args.get('user')
        if user and user == 'zhangsan':
            return original_function(*args, **kwargs)
        else:
            return '请先登录'

    return decorated_function


app.add_url_rule(rule='/page1', view_func=Page1.as_view('Page1'))
app.add_url_rule(rule='/page2', view_func=Page2.as_view('Page2'))

if __name__ == '__main__':
    app.run(debug=True)
