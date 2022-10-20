'''
逝者如斯夫, 不舍昼夜 -- 孔夫子
@Auhor    : Dohoo Zou
Project   : gitCode
FileName  : middleware.py
IDE       : PyCharm
CreateTime: 2022-10-20 11:17:54
'''

"""
中间件的作用：每次请求和响应的时候都会调用
中间件的使用：可以判断每次请求中cookie是否携带某些信息，如username之类的
多个中间件调用顺序：先入后出
"""


def simple_middleware(get_response):
    def middleware(request):
        print('调用中间件前')
        response = get_response(request)
        print('调用中间件后')
        return response

    return middleware
