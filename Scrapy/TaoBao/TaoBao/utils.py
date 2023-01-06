'''
逝者如斯夫, 不舍昼夜 -- 孔夫子
@Auhor    : Dohoo Zou
Project   : gitCode
FileName  : utils.py
IDE       : PyCharm
CreateTime: 2023-01-04 16:36:41
'''
import json
from selenium import webdriver


def create_chrome_driver(*, headless=False):
    """创建浏览器对象"""
    options = webdriver.ChromeOptions()
    # 浏览器是否有窗口
    if headless:
        options.add_argument('--headless')
    # 浏览器不显示测试软件控制
    options.add_experimental_option('excludeSwitches', ['enable-automation'])
    options.add_experimental_option('useAutomationExtension', False)
    browser = webdriver.Chrome(executable_path='/Users/dohozou/Downloads/Mac/chromedriver', options=options)
    # 防止selenium的反爬
    browser.execute_cdp_cmd(
        'Page.addScriptToEvaluateOnNewDocument',
        {'source': 'Object.defineProperty(navigator, "webdriver", {get: () => undefined})'}
    )
    return browser


def add_cookies(browser, cookie_file):
    with open(cookie_file, 'r') as file:
        cookie_list = json.load(file)
        for cookie_dict in cookie_list:
            if cookie_dict['secure']:
                browser.add_cookie(cookie_dict)
