#!usr/bin/env python
# -*- coding:utf-8 _*-
"""
# author: 小菠萝测试笔记
# blog:  https://www.cnblogs.com/poloyy/
# time: 2021/9/7 11:18 下午
# file: 18_实战6.py
"""


# 课程类
class Course(object):
    def __init__(self, name, price):
        # 课程名、课程价格：私有属性
        self.__name = name
        self.__price = price

    @property
    def name(self):
        return self.__name

    @name.setter
    def name(self, name):
        self.__name = name

    @property
    def price(self):
        return self.__price

    @price.setter
    def price(self, price):
        self.__price = price


# 人类
class Person(object):
    def __init__(self, name, sex, phone):
        self.name = name
        self.sex = sex
        self.phone = phone

    def __str__(self):
        return f"姓名：{self.name}, 性别{self.sex}, 电话：{self.phone}"


# 学生类
class Student(Person):
    def __init__(self, name, sex, phone, balance):
        super(Student, self).__init__(name, sex, phone)
        # 学生余额、报名的班级：私有属性
        self.__balance = balance
        self.__class_list = []

    @property
    def classList(self):
        return [class_.name for class_ in self.__class_list]

    # 报名班级
    def addClass(self, class_):
        price = class_.price
        # 1、如果学生余额大于班级费用
        if self.__balance > price:
            # 2、报名成功
            self.__class_list.append(class_)
            # 3、减去报名费
            self.__balance -= price
            # 4、班级的学生列表也需要添加当前学生
            class_.addStudent(self)
            # 5、班级总额增加
            class_.totalBalance()
            return
        print("余额不足，无法报名班级")

    # 退学
    def removeClass(self, class_):
        if class_ in self.__class_list:
            # 1、报名班级列表移除
            self.__class_list.remove(class_)
            # 2、班级学生列表移除当前学生
            class_.removeStudent(self)
        print("班级不存在，无法退学")


# 员工类
class Employ(Person):
    def __init__(self, name, sex, phone):
        super(Employ, self).__init__(name, sex, phone)
        # 工资：私有属性
        self.money = 0


# 老师类
class Teacher(Employ):
    def __init__(self, name, sex, phone):
        super(Teacher, self).__init__(name, sex, phone)
        # 授课班级：私有属性
        self.__class_list = []

    @property
    def classList(self):
        return [class_.name for class_ in self.__class_list]

    # 授课
    def teachClass(self, class_):
        # 1、授课列表添加班级
        self.__class_list.append(class_)
        # 2、班级添加授课老师
        class_.teacher = self


# 班级类
class Class(object):
    def __init__(self, name):
        # 班级名、班级费用、课程列表、学生类表、班级老师、所属学校：私有属性
        self.__name = name
        self.__price = 0
        self.__course_list = []
        self.__student_list = []
        self.__teacher = None
        self.__balance = 0
        self.__school = None

    @property
    def name(self):
        return self.__name

    @name.setter
    def name(self, name):
        self.__name = name

    @property
    def school(self):
        return self.__school.name

    @school.setter
    def school(self, school):
        self.__school = school

    @property
    def price(self):
        return self.__price

    @property
    def courseList(self):
        return self.__course_list

    def addCourse(self, course):
        # 1、班级费用累加课程费用
        self.__price += course.price
        # 2、添加到课程列表
        self.__course_list.append(course)

    @property
    def studentList(self):
        return [stu.name for stu in self.__student_list]

    def addStudent(self, student):
        self.__student_list.append(student)

    def removeStudent(self, student):
        self.__student_list.remove(student)

    @property
    def teacher(self):
        return self.__teacher

    @teacher.setter
    def teacher(self, teacher):
        self.__teacher = teacher

    @property
    def balance(self):
        return self.__balance

    # 统计班级一个班级收入
    def totalBalance(self):
        self.__balance = len(self.__student_list) * self.__price


# 学校类
class School(object):
    def __init__(self, name, balance):
        # 学校名、学校余额、学校员工列表：公共属性
        self.name = name
        self.balance = balance
        self.employ_list = []
        # 分校列表：私有属性
        self.__school_list = []

    def __str__(self):
        return f"学校：{self.name} 余额：{self.balance}"

    # 获取学校分校列表
    @property
    def schoolList(self):
        return [school.name for school in self.__school_list]

    # 添加分校
    def addBranchSchool(self, school):
        self.__school_list.append(school)

    # 添加员工
    def addEmploy(self, employ):
        self.employ_list.append(employ)

    # 查看员工列表
    def getEmploy(self):
        return [emp.name for emp in self.employ_list]

    # 统计各分校的账户余额
    def getTotalBalance(self):
        res = {}
        sum = 0
        for school in self.__school_list:
            # 1、结算一次分校余额
            school.getTotalBalance()
            res[school.name] = school.balance
            # 2、累加分校余额
            sum += school.balance
        res[self.name] = sum
        return res

    # 统计员工人数
    def getTotalEmploy(self):
        sum_emp = 0
        for school in self.__school_list:
            sum_emp += len(school.employ_list)
        sum_emp += len(self.employ_list)
        return sum_emp

    # 统计学生总人数
    def getTotalStudent(self):
        sum_stu = 0
        for school in self.__school_list:
            sum_stu += school.getTotalStudent()
        return sum_stu


# 分校类
class BranchSchool(School):
    def __init__(self, name, balance):
        super(BranchSchool, self).__init__(name, balance)
        # 分校的班级列表：私有属性
        self.__class_list = []

    # 获取班级列表
    @property
    def classList(self):
        return [class_.name for class_ in self.__class_list]

    # 添加班级
    def addClass(self, class_):
        # 1、添加班级
        self.__class_list.append(class_)
        # 2、添加老师员工
        self.addEmploy(class_.teacher)

    # 获取总的余额
    def getTotalBalance(self):
        for class_ in self.__class_list:
            # 1、结算班级收入
            class_.totalBalance()
            # 2、累加班级收入
            self.balance += class_.balance

    # 获取学生总人数
    def getTotalStudent(self):
        sum_stu = 0
        for class_ in self.__class_list:
            sum_stu += len(class_.studentList)
        return sum_stu


# 总校
school = School("小菠萝总校", 100000)
# 分校
bj1 = BranchSchool("小猿圈北京分校", 2222)
sz1 = BranchSchool("深圳南山大学城分校", 5555)

# 添加分校
school.addBranchSchool(bj1)
school.addBranchSchool(sz1)

# 初始化班级
class1 = Class("Python 基础班级")
class2 = Class("Python 进阶班级")

# 初始化课程
c1 = Course("Python 基础", 666)
c2 = Course("Python 进阶", 1666)
c3 = Course("Python 实战", 2666)

# 添加课程
class1.addCourse(c1)
class1.addCourse(c2)
class2.addCourse(c2)
class2.addCourse(c3)

# 初始化老师
tea1 = Teacher("小菠萝老师", "girl", 1355001232)
tea2 = Teacher("大菠萝老师", "boy", 1355001232)

# 老师授课
tea1.teachClass(class1)
tea2.teachClass(class2)

bj1.addClass(class1)
sz1.addClass(class2)

# 初始化学生
stu1 = Student("小菠萝", "girl", 1355001232, 50000)
stu2 = Student("大菠萝", "boy", 1355001231, 50000)
stu3 = Student("大大菠萝", "girl", 1355001233, 10000)
# 学生报名
stu1.addClass(class1)
stu1.addClass(class2)
stu2.addClass(class1)
stu3.addClass(class2)

# 普通员工
emp1 = Employ("小菠萝员工", "girl", 1355001232)
emp2 = Employ("大菠萝员工", "boy", 1355001231)
emp3 = Employ("小小菠萝员工", "girl", 1355001233)

print(bj1.getTotalStudent())
print(school.getTotalBalance())
print(school.getTotalEmploy())
