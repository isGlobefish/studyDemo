from django.db import models


# Create your models here.

class BookInfo(models.Model):
    name = models.CharField(max_length=10)

    def __str__(self):
        return self.name

    # 改变数据库名字
    class Meta:
        db_table = "book_tushu"



'''
补充知识：Django ForeignKey ondelete 级联操作
CASCADE：删除一并删除关联表下的所有的信息；
PROTECT：删除信息时，采取保护机制，抛出错误：即不删除关联表的内容；
SET_NULL：只有当null=True才将关联的内容置空；
SET_DEFAULT：设置为默认值；
SET()：括号里可以是函数，设置为自己定义的东西；
DO_NOTHING：字面的意思，啥也不干，你删除你的干我毛线关系
'''


class PersonInfo(models.Model):
    Name = models.CharField(max_length=10)
    gender = models.BooleanField()
    book = models.ForeignKey(BookInfo, on_delete=models.PROTECT)
