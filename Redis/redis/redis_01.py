'''
逝者如斯夫, 不舍昼夜 -- 孔夫子
@Auhor    : Dohoo Zou
Project   : gitCode
FileName  : redis_01.py
IDE       : PyCharm
CreateTime: 2022-10-13 08:15:53
'''
import redis

rd = redis.Redis(host='10.211.55.9', port=6378, password='00000000')
# rd.set('age', 19)
print(rd.get('age'))

