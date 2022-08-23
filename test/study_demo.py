import sys


class UserClass(object):

    def __init__(self):
        self._User = []

    def ___new__(cls, User=[]):
        if User != []:
            return None
        else:
            return object.__new__(cls)

    def __str__(self):
        return '__str__: 此时的_User为: %s' % self._User

    '''
    __repr__: 存在__str__时不生效, 自定义输出的实例化对象信息
    '''

    def __repr__(self):
        return '__repr__: 此时的_User为: %s' % self._User

    def __del__(self):
        print('结束程序,自动调用析构方法,释放内存'.center(53, '-'))

    # 新增
    def add_user(self):
        user_name = input('请输入用户名:')
        user_age = input('请输入年龄:')
        user_sex = input('请输入性别:')

        if user_name and user_age and user_sex:
            user_dict = {
                'user_name': user_name,
                'user_age': user_age,
                'user_sex': user_sex
            }
            self._User.append(user_dict)

    # 删除
    def delete_user(self):
        del_user_name = input('请输入删除用户的姓名:')
        if del_user_name in [user.get('user_name') for user in self._User]:
            for user_dict in self._User:
                if user_dict.get('user_name') == del_user_name:
                    self._User.remove(user_dict)
                    print(str(user_dict) + ' 已删除')
            print('此时的用户剩' + str(self._User))
        else:
            print('不存在该用户, 请重新输入')

    # 更新
    def update_user(self):
        update_user_name = input('请输入更新用户的姓名:')
        if update_user_name in [user.get('user_name') for user in self._User]:
            for user_dict in self._User:
                if user_dict.get('user_name') == update_user_name:
                    update_user_age = input('请输入更新用户的年龄(不填默认旧数据):')
                    update_user_sex = input('请输入更新用户的性别(不填默认旧数据):')
                    if update_user_age != user_dict.get('user_age') and update_user_age != '':
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
                    print(str(user_dict) + '更新成功')
        else:
            print('不存在该用户, 请重新输入')

    # 查询全部
    def query_all_user(self):
        for user in self._User:
            print(user)

    # 指定查询
    def query_one_user(self):
        qury_user_name = input('请输入要查询用户的姓名:')
        if qury_user_name in [user.get('user_name') for user in self._User]:
            for user_dict in self._User:
                if user_dict.get('user_name') == qury_user_name:
                    print(str(user_dict) + ' 查询成功')
        else:
            print('不存在该用户, 请重新输入')


class ChildClass(UserClass):
    pass


if __name__ == '__main__':
    # 实例化对象
    opts = UserClass()
    print(opts)

    while True:
        selectNum = int(input('(1.新增 2.删除 3.更新 4.查询ALL 5.查询one 6.中止)_'))

        if selectNum == 1:
            print('新增用户')
            opts.add_user()
        elif selectNum == 2:
            print('删除用户')
            opts.delete_user()
        elif selectNum == 3:
            print('更新用户')
            opts.update_user()
        elif selectNum == 4:
            print('查询所有的用户')
            opts.query_all_user()
        elif selectNum == 5:
            print('按照输入信息查询用户')
            opts.query_one_user()
        elif selectNum == 6:
            sys.exit()
