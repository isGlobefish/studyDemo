'''
逝者如斯夫, 不舍昼夜 -- 孔夫子
@Auhor    : Dohoo Zou
Project   : gitCode
FileName  : test2.py
IDE       : PyCharm
CreateTime: 2023-01-04 17:34:36
'''
from utils import create_chrome_driver, add_cookies

browser = create_chrome_driver()
browser.get('https://www.taobao.com')
add_cookies(browser, 'taobao.json')
browser.get('https://s.taobao.com/search?q=%E6%89%8B%E6%9C%BA&s=0')
