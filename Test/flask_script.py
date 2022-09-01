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


if __name__ == '__main__':
    app.run(debug=True)
