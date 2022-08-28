'''
逝者如斯夫, 不舍昼夜 -- 孔夫子
@Auhor    : Dohoo Zou
Project   : gitCode
FileName  : mysql_operate.py
IDE       : PyCharm
CreateTime: 2022-08-19 21:29:00
'''

import pymysql


class OperateDBClass(object):

    def __init__(self, host: str = 'localhost', port: int = 3306, username: str = 'root', password: str = '123456', database: str = None,
                 charset: str = 'utf-8') -> None:
        self._host = host
        self._port = port
        self._username = username
        self._pwd = password
        self._db = database
        self._charset = charset

    def connect_db(self):
        # 建立链接
        conn = pymysql.connect(
            host=self._host,
            port=self._port,
            user=self._username,
            password=self._password,
            database=self._db,
            charset=self._charset
        )

        # 获取游标
        cursor = conn.cursor()
        # 执行语句
        sql = 'select * from userinfo'
        res = cursor.execute(sql)
        # 关闭游标
        cursor.close()
        # 关闭连接
        conn.close()

    # 增加
    def insert(self):
        pass

    # 删除
    def delete(self):
        pass

    # 更新
    def update(self):
        pass

    # 查询
    def query(self):
        pass
