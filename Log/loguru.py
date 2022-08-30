'''
逝者如斯夫, 不舍昼夜 -- 孔夫子
@Auhor    : Dohoo Zou
Project   : gitCode
FileName  : loguru.py
IDE       : PyCharm
CreateTime: 2022-08-31 02:24:34
'''

import time
from loguru import logger
from pathlib import Path

loguru_path = Path.joinpath(Path.cwd(), "Log")
t = time.strftime("%Y-%m-%d")


class Loggings:

    __instance = None

    logger.add(f"{loguru_path}/loguru_{t}.log", rotation="500MB", encoding="utf-8", enqueue=True, retention="10 days")

    def __new__(cls, *args, **kwargs):
        if not cls.__instance:
            cls.__instance = super(Loggings, cls).__new__(cls, *args, **kwargs)
        return cls.__instance

    def info(self, msg):
        return logger.info(msg)

    def debug(self, msg):
        return logger.debug(msg)

    def warning(self, msg):
        return logger.warning(msg)

    def error(self, msg):
        return logger.error(msg)

    def critical(self, msg):
        return logger.critical(msg)


loggings = Loggings()

if __name__ == '__main__':
    loggings.info("中文test")
    loggings.debug("中文test")
    loggings.warning("中文test")
    loggings.error("中文test")
    loggings.critical("中文test")

    logger.info('If you are using Python {}, prefer {feature} of course!', 3.6, feature='f-strings')
    n1 = "cool"
    n2 = [1, 2, 3]
    logger.info(f'If you are using Python {n1}, prefer {n2} of course!')