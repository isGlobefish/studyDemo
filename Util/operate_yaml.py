'''
逝者如斯夫, 不舍昼夜 -- 孔夫子
@Auhor    : Dohoo Zou
Project   : gitCode
FileName  : operate_yaml.py
IDE       : PyCharm
CreateTime: 2022-08-31 02:44:53
'''
import yaml


# 单例设计模式
class YamlUtil:
    __instance = None

    def __new__(cls, *args, **kwargs):
        if not cls.__instance:
            print("YamlUtil first init")
            cls.__instance = super(YamlUtil, cls).__new__(cls, *args, **kwargs)
        return cls.__instance

    def read_yaml(self, path):
        with open(path, encoding="utf-8") as f:
            result = f.read()
            result = yaml.load(result, Loader=yaml.FullLoader)
            return result

    def write_yaml(self, path, data):
        with open(path, "w", encoding="utf-8") as f:
            yaml.dump(data, f, Dumper=yaml.SafeDumper)


yamlUtil = YamlUtil()

yamlUtil.read_yaml('/Users/dohozou/Desktop/Code/gitCode/Config/mysql.yaml')
