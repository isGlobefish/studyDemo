'''
Mysqldb 不兼容 python3.5 以后的版本
解决办法：
使用pymysql代替MySQLdb
'''
import pymysql

pymysql.version_info = (1, 4, 13, "final", 0)
pymysql.install_as_MySQLdb()