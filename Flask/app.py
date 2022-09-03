'''
逝者如斯夫, 不舍昼夜 -- 孔夫子
@Auhor    : Dohoo Zou
Project   : gitCode
FileName  : app.py
IDE       : PyCharm
CreateTime: 2022-09-01 00:08:02
'''

from flask import Flask, request, render_template, Response

app = Flask(__name__)


@app.route('/hello')
def hell_word():
    return '<b>hello word</b>'


@app.route('/get', methods=['GET'])
def get_():
    return '这是一个get方法'


@app.route('/post', methods=['POST'])
def post_():
    return dict(message='这是一个post请求')


@app.route('/delandput', methods=['DELETE', 'PUT'])
def del_put():
    return {'mesage': ['delete', 'put']}


def echo(key, value):
    print('%-10s = %s' % (key, value))


@app.route('/query')
def query():
    echo('url', request.url)
    echo('basr_url', request.base_url)
    echo('host_url', request.host_url)
    echo('path', request.path)
    echo('full_path', request.full_path)

    print(request.args)
    print('userId = %s' % request.args['userId'])
    return 'hello url'


@app.route('/user/<string:user_name>')
def show_user_name(user_name):
    return 'user_name = %s' % user_name


@app.route('/user/<int:user_age>')
def show_user_age(user_age):
    return 'user_age = %s' % user_age


@app.route('/user/<float:user_price>')
def show_user_price(user_price):
    return 'user_price = %s' % user_price


@app.route('/user/<path:user_path>')
def show_user_path(user_path):
    return 'user_path = %s' % user_path


# 默认GET，postman 发起 GET 请求，params url里面传数据
@app.route('/method_args')
def method_args():
    return {'name': request.args['name'], 'age': request.args['age'], 'method': request.args}


# postman 发起 POST 请求，form-data 传数据
@app.route('/method_form', methods=['POST'])
def method_form():
    return {'name': request.form['name'], 'age': request.form['age'], 'method': request.form}


# postman 发起 POST 请求，raw -> json 传数据
@app.route('/method_json', methods=['POST'])
def method_json():
    return {'name': request.json['name'], 'age': request.json['age'], 'method': request.json}


# postman 发起 POST 请求，form-data 传数据
@app.route('/method_values', methods=['POST'])
def method_values():
    return {'name': request.values['name'], 'age': request.values['age'], 'method': request.values}


@app.route('/index1')
def index():
    return render_template('index.html', name='tom', age=10)


@app.route('/')
def index2():
    return render_template('index2.html', string='hello word', lists=[_ for _ in range(1, 5)],
                           dict=dict(name='china'.capitalize(), age=120))


# 导入蓝图类
from Flask.news import news
from Flask.products import products

# 注册蓝图
app.register_blueprint(news.blueprint)
app.register_blueprint(products.blueprint)

# 标准类视图
from flask import views
from flask.typing import ResponseReturnValue


class view_test1(views.View):

    def dispatch_request(self) -> ResponseReturnValue:
        return 'hello 视图标准类'


class view_test2(views.View):

    def dispatch_request(self) -> ResponseReturnValue:
        return {"msg": "success", "code": 0}

    @staticmethod
    def as_view(name, **kwargs):
        view = view_test2()
        return view.dispatch_request


# 将路由规则 / 和视图类 view_test 进行绑定
app.add_url_rule(rule='/view1', view_func=view_test1.as_view('view1'))
app.add_url_rule(rule="/view2", view_func=view_test2.as_view("view2"))

# 继承视图类
from Flask.viewclass.baseview import BaseView


class UserView(BaseView):

    def get_template(self):
        return 'userview.html'

    def get_data(self):
        return {
            'name': '邹德豪',
            'gender': '男',
            'age': 18
        }


app.add_url_rule('/userview', view_func=UserView.as_view('UserView'))

from flask import Flask, request, views
from functools import wraps


# 定义修饰器
def check_login(original_function):
    @wraps(original_function)
    def decorator_function(*args, **kwargs):
        user = request.args.get('user')
        if user and user == 'zhangsan':
            return original_function(*args, **kwargs)
        else:
            return '请先登录！！！'

    return decorator_function


@app.route('/decorator')
@check_login
def decorator():
    return '修饰器'


class Decorator(views.View):
    decorators = [check_login]

    def dispatch_request(self) -> ResponseReturnValue:
        return '修饰器在视图类中使用'


app.add_url_rule('/decorator_view', view_func=Decorator.as_view('decorator_view'))


# 获取 cookie
@app.route('/get_cookie')
def get_cooke():
    cookie = request.cookies.get('poloyy')
    return render_template('get_cookie.html', cookie=cookie)


# 设置 cookie
@app.route('/set_cookie')
def set_cooke():
    html = render_template('js_cookie.html')
    response = Response(html)
    response.set_cookie('poloyy', 'https://www.cnblogs.com/poloyy')
    return response


# 删除 cookie
@app.route("/del_cookie")
def del_cookie():
    html = render_template("js_cookie.html")
    response = Response(html)
    response.delete_cookie("poloyy")
    return response


from flask import session
import os

app.config['SECRET_KEY'] = os.urandom(24)


# 设置 session
@app.route('/set_session')
def set_session():
    session['user'] = 'poloyy'
    session['pwd'] = 'passwoed'
    return render_template('session_query.html', user=session.get('user'), pwd=session.get('pwd'))


# 获取 session
@app.route('/get_session')
def get_session():
    return render_template('session_query.html', user=session.get('user'), pwd=session.get('pwd'))


# 删除 session
@app.route('/del_session')
def del_session():
    session.pop('user')
    return render_template('session_query.html', user=session.get('user'), pwd=session.get('pwd'))


# 清空 session
@app.route('/clear_session')
def clear_session():
    session.clear()
    return render_template('session_query.html', user=session.get('user'), pwd=session.get('pwd'))


if __name__ == '__main__':
    # 默认主机: 127.0.0.1 端口: 5000
    app.run(host='127.0.0.1', port=5000, debug=True)
