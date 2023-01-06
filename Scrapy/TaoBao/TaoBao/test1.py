'''
逝者如斯夫, 不舍昼夜 -- 孔夫子
@Auhor    : Dohoo Zou
Project   : gitCode
FileName  : test1.py
IDE       : PyCharm
CreateTime: 2023-01-04 17:34:50
'''
import json
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.wait import WebDriverWait
from utils import create_chrome_driver

browser = create_chrome_driver()
# browser.get('https://login.taobao.com')
browser.get('https://s.taobao.com/search?q=手机&s=0')
# 隐式等待
browser.implicitly_wait(5)
username_input = browser.find_element(By.XPATH, '/html/body/div/div[2]/div[3]/div/div/div/div[2]/div/form/div[1]/div[2]/input')
username_input.send_keys('13267854059')
password_input = browser.find_element(By.XPATH, '/html/body/div/div[2]/div[3]/div/div/div/div[2]/div/form/div[2]/div[2]/input')
password_input.send_keys('zdh19970516')
login_button = browser.find_element(By.XPATH, '/html/body/div/div[2]/div[3]/div/div/div/div[2]/div/form/div[4]/button')
login_button.click()
# 显式等待
# wait_obj = WebDriverWait(browser, 10)
# wait_obj.until(expected_conditions.presence_of_element_located((By.ID, 'kw1')))
time.sleep(5)
# 获取Cookie数据写入文件
with open('taobao.json', 'w') as file:
    json.dump(browser.get_cookies(), file)
