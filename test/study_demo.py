import json
import sys


class UserClass(object):

    def __init__(self):
        self._USERLIST = []
        self._SEXLIST = ['1', '0', 'Man', 'Woman', 'Male', 'Female', '男', '女']

    def __new__(cls, _USERLIST=[]):
        if _USERLIST != []:
            return None
        else:
            return object.__new__(cls)

    def __str__(self):
        return '__str__: 此时的_USERLIST为: %s' % self._USERLIST

    '''
    __repr__: 存在__str__时不生效, 自定义输出的实例化对象信息
    '''

    def __repr__(self):
        return '__repr__: 此时的_USERLIST为: %s' % self._USERLIST

    def __del__(self):
        print('程序终止,自动调用析构方法,释放内存'.center(50, '-'))

    # 新增
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

    # 删除
    def delete_user(self):
        del_user_name = input('请输入删除用户的姓名:')
        if del_user_name in [user.get('user_name') for user in self._USERLIST]:
            for user_dict in self._USERLIST:
                if user_dict.get('user_name') == del_user_name:
                    self._USERLIST.remove(user_dict)
                    print(str(user_dict) + ' 已删除')
            print(f'此时的用户剩{self._USERLIST}')
        else:
            print('不存在该用户, 请重新输入')

    # 更新
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

    # 查询全部
    def query_all_user(self):
        for order, user in enumerate(self._USERLIST, start=1):
            print('用户%d: %s' % (order, user))

    # 指定查询
    def query_one_user(self):
        qury_user_name = input('请输入要查询用户的姓名:')
        if qury_user_name in [user.get('user_name') for user in self._USERLIST]:
            for user_dict in self._USERLIST:
                if user_dict.get('user_name') == qury_user_name:
                    print('%s 查询成功' % str(json.dumps(user_dict, indent=4, ensure_ascii=False)))
        else:
            print('不存在该用户, 请重新输入')


class PrivateClass(object):
    pass


class ExtendClass(object):
    pass


class ChildClass(UserClass, PrivateClass, ExtendClass):
    pass


if __name__ == '__main__':
    # 实例化对象
    opts = ChildClass()
    print('{0} \nMRO方法搜索顺序: {1}'.format(opts, ChildClass.__mro__))

    while True:
        selectNum = int(input('(1.新增 2.删除 3.更新 4.查询ALL 5.查询one 6.中止)_'))

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
        elif selectNum == 6:
            sys.exit()
