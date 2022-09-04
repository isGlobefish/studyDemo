'''
逝者如斯夫, 不舍昼夜 -- 孔夫子
@Auhor    : Dohoo Zou
Project   : gitCode
FileName  : study_demo.py
IDE       : PyCharm
CreateTime: 2022-08-18 21:29:00
'''

import json
import sys
from typing import Callable
from random import randint
from threading import Thread
from time import time, sleep


# 用户类
class UserClass(object):

    def __init__(self):
        self._USERLIST = []
        self._SEXLIST = ['1', '0', 'Man', 'Woman', 'Male', 'Female', '男', '女']

    '''
    __new__ 主要作用
    在内存中为实例对象分配空间
    返回对象的引用给 Python 解释器
    Python 的解释器获得对象的引用后，将对象的引用作为第一个参数，传递给 __init__ 方法
    
    重写 __new__ 方法
    重写的代码是固定的
    重写 __new__ 方法一定要在最后 return super().__new__(cls) 
    如果不 return（像上面代码栗子一样），Python 的解释器得不到分配了空间的对象引用，就不会调用对象的初始化方法（__init__）
    重点：__new__ 是一个静态方法，在调用时需要主动传递 cls 参
    '''

    def __new__(cls, _USERLIST=[]):
        # 如果类属性 is None，则调用父类方法分配内存空间，并赋值给类属性
        if _USERLIST == []:
            return object.__new__(cls)
            # 如果类属性已有对象引用，则直接返回
        else:
            return None

    def __str__(self):
        return '__str__: 此时的_USERLIST为: %s' % self._USERLIST

    '''
    __repr__: 存在__str__时不生效, 自定义输出的实例化对象信息
    '''

    def __repr__(self):
        return '__repr__: 此时的_USERLIST为: %s' % self._USERLIST

    def __del__(self):
        print('程序终止,自动调用析构方法,释放内存'.center(50, '-'))

    # 初始化用户列表_USERLIST
    def reStart(self):
        self._USERLIST = []
        return self._USERLIST

    # 1. 新增
    def add_user(self):
        user_name = input('请输入用户姓名:')
        if user_name in [name.get('user_name') for name in self._USERLIST]:
            isExist = input('已存在该用户, 是否替换Y/N:').upper()
            if isExist == 'Y':
                for user_dict in self._USERLIST:
                    if user_dict.get('user_name') == user_name:
                        self._USERLIST.remove(user_dict)
        user_age = (lambda age: int(age) if age.isdigit() > 0 else '输入的年龄有误！！！')(input('请输入用户年龄:'))
        user_sex = (lambda x: x if x in self._SEXLIST else '性别仅有以下情况:' + str(self._SEXLIST))((input('请输入用户性别:')).capitalize())
        if user_name and user_age and user_sex:
            user_dict = {
                'user_name': user_name,
                'user_age': user_age,
                'user_sex': user_sex
            }
            self._USERLIST.append(user_dict)
        else:
            print('用户姓名or年龄or性别为空, 请重新输入')
        self.query_all_user()

    # 2. 删除
    def delete_user(self):
        del_user_name = input('请输入删除用户的姓名(不输入默认清空全部):')
        if del_user_name in [user.get('user_name') for user in self._USERLIST] and del_user_name != '':
            for user_dict in self._USERLIST:
                if user_dict.get('user_name') == del_user_name:
                    self._USERLIST.remove(user_dict)
                    print(str(user_dict) + ' 已删除')
            print(f'此时的用户剩{self._USERLIST}')
        elif del_user_name not in [user.get('user_name') for user in self._USERLIST] and del_user_name == '':
            self.reStart()
            print('用户已清空！！！')
            self.query_all_user()
        else:
            print('不存在该用户, 请重新输入')

    # 3. 更新
    def update_user(self):
        update_user_name = input('请输入更新用户的姓名:')
        if update_user_name in [user.get('user_name') for user in self._USERLIST]:
            for user_dict in self._USERLIST:
                if user_dict.get('user_name') == update_user_name:
                    update_user_age = (lambda age: int(age) if age.isdigit() > 0 else 0)(input('请输入更新用户的年龄(不填默认旧数据):'))
                    update_user_sex = (lambda sex: sex if sex in self._SEXLIST else '')((input('请输入更新用户的性别(不填默认旧数据):')).capitalize())
                    if update_user_age != user_dict.get('user_age') and update_user_age != 0:
                        user_dict.update({
                            'user_age': update_user_age
                        })
                    else:
                        print('用户年龄不变')
                    if update_user_sex != user_dict.get('user_sex') and update_user_sex != '':
                        user_dict.update({
                            'user_sex': update_user_sex
                        })
                    else:
                        print('用户年龄不变')
                    print(str(user_dict) + ' 更新成功')
        else:
            print('不存在该用户, 请重新输入')

    # 4. 查询全部
    def query_all_user(self):
        if self._USERLIST != None and self._USERLIST != []:
            for order, user in enumerate(self._USERLIST, start=1):
                print('用户%d: %s' % (order, user))
        else:
            print(f'用户数据__USERLIST: {self._USERLIST}')

    # 5. 指定查询
    def query_one_user(self):
        qury_user_name = input('请输入要查询用户的姓名:')
        if qury_user_name in [user.get('user_name') for user in self._USERLIST]:
            for user_dict in self._USERLIST:
                if user_dict.get('user_name') == qury_user_name:
                    print('%s 查询成功' % str(json.dumps(user_dict, indent=4, ensure_ascii=False)))
        else:
            print('不存在该用户, 请重新输入')


# 类内部方法
class PrivateClass(object):
    # 公共属性/变量
    SUM = 0.1
    # 保护属性/变量
    _SUM = 0.2
    # 私有属性/变量
    __SUM = 0.3

    # 构造方法
    def __init__(self, name, age, sex):
        self.name = name
        self._age = age
        self.__sex = sex

    # 实例方法
    def getName(self):
        print(f'实例方法的name: {self.name}, _age: {self._age}, __sex: {self.__sex}')

    # 受保护方法
    def _getName(self):
        print(f'保护方法的name: {self.name}, _age: {self._age}, __sex: {self.__sex}')

    # 私有方法
    def __getName(self):
        print(f'私有方法的name: {self.name}, _age: {self._age}, __sex: {self.__sex}')

    # 类方法
    @classmethod
    def sum_01(cls, sum):
        cls.SUM += sum
        cls._SUM += sum
        cls.__SUM += sum
        print('类方法: SUM的值为 %f, SUM的值为 %f, SUM的值为 %f' % (cls.SUM, cls._SUM, cls.__SUM))

    # 静态方法
    @staticmethod
    def sum_02(value1, value2):
        print('静态方法: {0} + {1} = {2}'.format(value1, value2, value1 + value2))


# 提供对外方法
class ExtendClass(object):

    def __init__(self, name, age, sex):
        self.name = name
        self._age = age
        self.__sex = sex

    @property
    def package_name(self):
        return self.name

    @package_name.setter
    def package_name(self, name):
        self.name = name

    @package_name.deleter
    def package_name(self):
        print(f'删除package_name了')

    @property
    def _package_age(self):
        return self._age

    @_package_age.setter
    def _package_age(self, age):
        self._age = age

    @_package_age.deleter
    def _package_age(self):
        print(f'删除_package_age了')

    @property
    def __package_sex(self):
        return self.__sex

    @__package_sex.setter
    def __package_sex(self, sex):
        self.__sex = sex

    @__package_sex.deleter
    def __package_sex(self):
        print(f'删除__package_sex了')


# 多继承
class ChildClass(UserClass, PrivateClass, ExtendClass):

    def __init__(self, child_name: str = '吴丽婷', child_age: int = 18, child_sex: bool = 0) -> None:
        self.child_name = child_name
        self.child_age = child_age
        self.child_sex = child_sex
        UserClass.__init__(self)
        PrivateClass.__init__(self, name=child_name, age=child_age, sex=child_sex)
        ExtendClass.__init__(self, name=child_name, age=child_age, sex=child_sex)


class XiaoWu(object):

    def __init__(self, name: str, age: int) -> None:
        self.name = name
        self.age = age

    def ta(self):
        print(f'{self.name} ta今年 {self.age}')


class XiaoZou(XiaoWu):

    def ta(self):
        print(f'{self.name} ta今年 {self.age}')


class TaSex:

    def __init__(self, sex: bool) -> None:
        self.sex = sex

    def taSex(self, xiao):
        print(f'ta的性别是 {self.sex}')
        xiao.ta()


# __call__: 使得类实例对象可以像普通函数那样被调用
class CallClass(object):

    def __init__(self, name: str) -> None:
        self.name = name

    def __call__(self, *args, **kwargs):
        print(self.name)
        print(args)
        print(kwargs)


# 多线程
class ThreadClass(Thread):
    __NUMBER = 0
    __SECONDS = 0

    def __init__(self, filename):
        super(ThreadClass, self).__init__()
        self.filename = filename

    @classmethod
    def add_number(cls):
        cls.__NUMBER += 1
        return cls.__NUMBER

    @classmethod
    def sum_seconds(cls, seconds):
        cls.__SECONDS += seconds
        return cls.__SECONDS

    def run(self) -> None:
        print(f'第{self.add_number()}个文件开始下载 {self.filename}')
        down_time = randint(5, 10)
        self.sum_seconds(down_time)
        sleep(down_time)
        print('%s 下载完成, 耗费了 %d 秒:' % (self.filename, down_time))


def run_thread():
    start = time()
    p1 = ThreadClass('小吴记仇本.pdf')
    p2 = ThreadClass('葵花宝典.pdf')
    p1.start()
    p2.start()
    p1.join()
    p2.join()
    end = time()
    print('总共耗时了 %.2f 秒 << 单线程所需耗时 %.2f 秒' % ((end - start), p2._ThreadClass__SECONDS))


if __name__ == '__main__':
    # 实例化对象
    opts = ChildClass()
    # 输出实例化之后的信息
    print('{0} \nMRO方法搜索顺序: {1}'.format(opts, ChildClass.__mro__))

    # todo
    while True:
        selectNum = (lambda inp: int(inp) if inp.isdigit() else 0)(
            input('「1.新增 2.删除 3.更新 4.查询ALL 5.查询one 6.中止 7.类内部 8.封装 9.多态 10.call 11.魔法函数 12.多线程」_'))

        if selectNum == 1:
            print('1. 新增用户')
            opts.add_user()
        elif selectNum == 2:
            print('2. 删除用户')
            opts.delete_user()
        elif selectNum == 3:
            print('3. 更新用户')
            opts.update_user()
        elif selectNum == 4:
            print('4. 查询所有的用户')
            opts.query_all_user()
        elif selectNum == 5:
            print('5. 按照输入信息查询用户')
            opts.query_one_user()
        elif selectNum == 6 or selectNum == 0:
            sys.exit()
        elif selectNum == 7:
            print('7. 类内部方法')
            print(f'公共属性SUM={opts.SUM}, _SUM={opts._SUM}, __SUM={opts._PrivateClass__SUM}')
            opts.getName()
            opts._getName()
            opts._PrivateClass__getName()
            opts.sum_01(sum=100)
            opts.sum_02(value1=1000, value2=1)
        elif selectNum == 8:
            print('8. 封装')
            print(
                f'封装(未修改)的package_name: {opts.package_name}, _package_age: {opts._package_age}, __package_sex: {opts._ExtendClass__package_sex}')
            opts.package_name, opts._package_age, opts._ExtendClass__package_sex = '邹德豪', 20, 1
            print(
                f'封装(修改后)的package_name: {opts.package_name}, _package_age: {opts._package_age}, __package_sex: {opts._ExtendClass__package_sex}')
            del opts.package_name, opts._package_age, opts._ExtendClass__package_sex
        elif selectNum == 9:
            print('9. 多态')
            xiaowu = XiaoWu(name='吴丽婷', age=18)
            xiaozou = XiaoZou(name='邹德豪', age=20)
            tasex = TaSex(sex=0).taSex(xiaowu)
            tasex = TaSex(sex=1).taSex(xiaozou)
        elif selectNum == 10:
            print('10. __call__')
            call = CallClass('乌拉圭')
            call(1, 2, 3, age=24, sex="girl")
            print(isinstance(call, Callable))
        elif selectNum == 11:
            print('11. 常见的魔法函数')
            print('11.1. __new__: https://www.cnblogs.com/poloyy/p/15236309.html')
            print('11.2. __init__: https://www.cnblogs.com/poloyy/p/15189562.html')
            print('11.3. __str__: https://www.cnblogs.com/poloyy/p/15202541.html')
            print('11.4. __repr__: https://www.cnblogs.com/poloyy/p/15250973.html')
            print('11.5. __del__: https://www.cnblogs.com/poloyy/p/15192098.html')
            print('11.6. __call__: https://www.cnblogs.com/poloyy/p/15253366.html')
            print('11.7. 全部魔法函数教程: https://www.cnblogs.com/poloyy/p/15245172.html')
        elif selectNum == 12:
            print('12. 多线程')
            run_thread()
