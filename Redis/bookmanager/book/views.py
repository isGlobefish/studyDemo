from django.shortcuts import render, redirect
from django.http import HttpResponse, HttpRequest
from book.models import BookInfo
import json

from django.urls import reverse


# 首页
def index(request):
    books = BookInfo.objects.all()
    # name = '元芳'
    # context = {'name': name}
    # return render(request, 'index.html', locals())
    # return render(request, 'index.html', context)
    context = {
        'books': books
    }
    # print(json.dumps(booklist))
    return render(request, 'index.html', context)
    # return HttpResponse('index')


#  重定向
def reindex(request):
    path = reverse('book:index')
    print(path)
    return redirect(path)


# 提取url特定的位置
def specil_url(request, category_id, book_id):
    return HttpResponse(f'第{category_id}页 / {book_id}个')


# 查询字符串
def query_string(request):
    params = request.GET
    # user = params['user']
    # user = params.get('user')
    user = params.getlist('user')
    # password = params['password']
    # password = params.get('password')
    password = params.get('password')
    return HttpResponse('用户: {}, 密码: {}'.format(user, password))


# 请求体（body）post数据
def body_post(request):
    # Post 表单数据
    # params = request.POST
    # user = params.get('user')
    # password = params.get('password')

    # Post json
    # 注意：json需要使用双引号
    # {
    #    "user": "zoudehao",
    #     "password": "123456"
    # }
    json_str = request.body.decode()  # jsom形式的字符串
    # json.loads() # 将json字符串转json
    # json.dumps() # 将json转json字符串
    data = json.loads(json_str)
    user = data.get("user")
    password = data.get("password")

    # Post 请求头
    # content_type = request.META['CONTENT_TYPE']

    """
    data = {'name': 'itcast'}
    # content: 传递字符串，不要传递对象、字典等
    # status: 100 - 599
    # content_type: 是一个MIME类型, 格式为大类/小类，如text/css, text/html, text/javascript
                                    application/json, image/png, image/gif, image/jpeg
    return HttpResponse(data, status=400, content_type='')
    
    传数据非要字符串太麻烦了，JsonResponse
    from django.http import JsonResponse
    return JsonResponse(data)
    """
    return HttpResponse('用户: {}, 密码: {}'.format(user, password))


'''
在客户端存储信息使用 Cookie
流程（原理）
第一次请求过程：
1. 我们浏览器第一次请求服务器的时候，不会携带任何cookie信息
2. 服务器接收到请求之后，发现 请求中没有任何cookie信息
3. 服务器设置一个cookie，这个cookie设置在响应中
4. 我们的浏览器接收到这个响应之后，发现响应中有cookie信息，浏览器将cookie信息保存起来

第二次及其之后的过程
5. 当我们的浏览器第二次及其之后的请求都会携带cookie信息
6. 我们的服务器接收到请求之后，会发现请求中携带的cookie信息，这样的话就认识是谁的请求了
'''


# 设置cookie
def set_cookie(request):
    # 先假设是第一次没有cookie信息
    # 获取username
    username = request.GET.get('username')
    # 在服务器中设置cookies
    response = HttpResponse('set_cookie')
    # max_age单位是秒，删除cookie的本质是max_age设置为0
    response.set_cookie('username', username, max_age=3600)
    return response


# 删除cookie的2种方法
# response.delete_cookie(key)
# responese.set_cookie(key, value, max_age=0)

# 获取cookie
def get_cookie(request):
    username = request.COOKIES.get('username')

    return HttpResponse(username)


'''
在服务端存储信息使用 Session
流程（原理）
第一次请求过程：
1. 我们第一次请求的时候可以携带一些信息（用户名/密码）cookie中没有任何信息
2. 当我们的服务器接收到这个请求之后，进行用户名和密码的验证，验证没有问题可以设置session信息
3. 在设置session信息的同时（session信息保存在服务器端），服务器会在相应头中设置一个session信息
4. 客户端（浏览器）在接收到响应之后，会将cookie信息保存起来（保存 session的信息）

第二次及其之后的过程
5. 第二次及其之后的请求都会携带session信息
6. 当服务器接收到这个请求之后，会获取到sessionid信息，然后进行验证，验证成功，则可以获取session信息（session信息保存在服务器端）
'''


def set_session(request):
    print(request.COOKIES)
    user_id = 6666
    request.session['user_id'] = user_id
    return HttpResponse(request.COOKIES)


def get_session(request):
    userid = request.session['user_id']
    return HttpResponse(userid)
