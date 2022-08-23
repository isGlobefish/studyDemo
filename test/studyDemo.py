import sys


class PersonClass(object):

    # 新增
    def add_person(self):
        pass

    # 删除

    # 更新

    # 查询全部

    # 指定查询


class ChildClass(PersonClass):
    pass


if __name__ == '__main__':
    # 实例化对象
    opts = PersonClass()

    selectNum = int(input('1.新增 2.删除 3.更新 4.查询ALL 5.查询one 6.中止'))

    if selectNum == 1:
        pass
