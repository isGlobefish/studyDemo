# -*- coding: utf-8 -*-
'''
@createTool    : PyCharm-2020.3.2
@projectName   : pythonProjectPy3.9
@originalAuthor: Made in win10.Sys design by deHao.Zou
@createTime    : 2021/01/21 16:20
'''


def DL():  # 大连
    print('\n' + "开始导出大连数据ing......")
    global sf
    global all_data1
    global driver
    sf = ['大连']
    all_data1 = []
    wei_zhi = "D:\\FilesCenter\\大客户数据\\HW\\"  # 大连下载路径
    options = webdriver.ChromeOptions()  # 打开设置
    prefs = {'profile.default_content_settings.popups': 0, 'download.default_directory': wei_zhi}  # 设置路径
    options.add_experimental_option('prefs', prefs)
    driver = webdriver.Chrome(executable_path='C:\\Program Files (x86)\\Google\\Chrome\\Application\\chromedriver.exe',
                              options=options)  # 打开浏览器
    driver.implicitly_wait(5)  # 隐式等待
    with open('D://FilesCenter//大客户数据//anquan.txt', 'r', encoding='utf8') as f:
        listCookies = json.loads(f.read())  # 读取cookies
    driver.get('http://srm.nepstar.cn')  # 进入登录界面
    driver.delete_all_cookies()  # 删旧cookies
    for cookie in listCookies:
        if 'expiry' in cookie:
            del cookie['expiry']
        driver.add_cookie(cookie)  # 新增cookies
    time.sleep(3)
    driver.get(
        'http://srm.nepstar.cn/ELSServer_HWXC/default2.jsp?account=110829_1001&loginChage=N&telphone1=18279409642')
    # 读取完cookie刷新页面
    time.sleep(1)
    button_a('//*[@id="treeMenu"]/li[7]/a')  # 数据查询
    time.sleep(0.5)
    button_a('//*[@id="salesOutSourcingInfoManage"]')  # 销售查询
    time.sleep(0.5)
    xf = driver.find_element_by_xpath('/html/body/div[2]/div/nav[2]/div[3]/iframe')  # 先通过xpath定位到iframe
    driver.switch_to.frame(xf)  # 再将定位对象传给switch_to.frame()方法
    button_a('/html/body/div[1]/div[2]/form/div[1]/div/div/span')  # 商品编码
    time.sleep(5)
    driver.switch_to.parent_frame()
    xf = driver.find_element_by_xpath('/html/body/div[4]/div/table/tbody/tr[2]/td/div/iframe')
    driver.switch_to.frame(xf)  # 切换
    button_a('/html/body/div/main/div[1]/div[1]/div[1]/table/thead/tr/th[2]/div/span/input')  # 全选
    time.sleep(0.5)
    button_a('/html/body/div/main/div[2]/button[1]')  # 确定
    time.sleep(0.5)
    driver.switch_to.parent_frame()
    xf = driver.find_element_by_xpath('/html/body/div[2]/div/nav[2]/div[3]/iframe')
    driver.switch_to.frame(xf)  # 切换
    input_b('/html/body/div[1]/div[2]/form/div[2]/div/div/input', time0)  # 查询日期
    time.sleep(0.5)
    input_b('/html/body/div[1]/div[2]/form/div[3]/div/div/input', time1)
    time.sleep(0.5)
    button_a('/html/body/div[1]/div[2]/form/div[4]/div/div/div/p')  # 联采合同
    time.sleep(0.5)
    button_a('/html/body/div[1]/div[2]/form/div[4]/div/div/div/div/ul/li[3]')
    time.sleep(0.5)
    button_a('/html/body/div[1]/div[2]/form/div[5]/div/div/div/p')  # 选地区
    time.sleep(0.5)
    ix = '/html/body/div[1]/div[2]/form/div[5]/div/div/div/div/ul/li[5]'  # 大连
    button_a(ix)  # 选好地区
    time.sleep(0.5)
    button_a('/html/body/div[1]/div[2]/form/button[2]')  # 查询
    time.sleep(7)
    button_a('/html/body/div[1]/div[3]/div[1]/div[1]/table/thead/tr/th[2]/div/span')  # 全选
    time.sleep(1)
    button_a('/html/body/div[1]/div[1]/nav/div/div/ul/li[1]/a')  # 查看明细
    time.sleep(1)
    driver.switch_to.parent_frame()  # 回退
    try:
        driver.find_element_by_xpath('/html/body/div[2]/div/nav[1]/ul/li[4]/a')  # 查看明细是否存在
        xf = driver.find_element_by_xpath('/html/body/div[2]/div/nav[2]/div[4]/iframe')  # 切入明细
        driver.switch_to.frame(xf)
        time.sleep(8)
        button_a('/html/body/div/div[1]/nav/div/div/ul/li[2]/a')  # 导出
        driver.switch_to.parent_frame()  # 切出
        time.sleep(6)
        element = driver.find_element_by_xpath('/html/body/div[2]/div/nav[1]/ul/li[4]/a/i[2]')
        ActionChains(driver).move_to_element(element).perform()  # 鼠标悬浮
        button_a('/html/body/div[2]/div/nav[1]/ul/li[4]/a/i[2]')  # 关闭明细
        time.sleep(8)
        all_excel = get_exce(wei_zhi)
        # 得到要合并的所有excel表格数据
        if (all_excel == 0):
            cprint("大连下载出错！！！", 'magenta', attrs=['bold', 'reverse', 'blink'])
        else:
            for excel in all_excel:
                fh = xlrd.open_workbook(excel)
                # 打开文件
                sheets = get_sheet(fh)
                # 获取文件下的sheet数量
                for sheet in range(len(sheets)):
                    row = get_sheetrow_num(sheets[sheet])
                    # 获取一个sheet下的所有的数据的行数
                    all_data1 = get_sheet_data(sheets[sheet], row, 0)
                    os.remove(excel)
        print("导出完成！ 大连")
        xf = driver.find_element_by_xpath('/html/body/div[2]/div/nav[2]/div[3]/iframe')  # 切入查询
        time.sleep(1)
        driver.switch_to.frame(xf)  # 切换
        button_a('/html/body/div[1]/div[3]/div[1]/div[1]/table/thead/tr/th[2]/div/span')  # 全不选
    except:
        cprint("大连->列无数据", 'magenta', attrs=['bold', 'reverse', 'blink'])
        xf = driver.find_element_by_xpath('/html/body/div[2]/div/nav[2]/div[3]/iframe')  # 切入查询
        driver.switch_to.frame(xf)  # 切换
        button_a('/html/body/div[1]/div[3]/div[1]/div[1]/table/thead/tr/th[2]/div/span')  # 全不选

    dictCookies = driver.get_cookies()  # 获取cookies
    jsonCookies = json.dumps(dictCookies)
    with open('D://FilesCenter//大客户数据//anquan.txt', 'w') as f:
        f.write(jsonCookies)  # 保存新cookies

    biao_tou = ['类别', '商品SAP编码', '商品名称', '规格', '单位', '店号/区域ID', '店名/区域', '销量', '过账日期', '合同价', '省份']
    all_data1.insert(0, biao_tou)  # 表头写入
    # 下面开始文件数据的写入
    new_excel = 'D:/FilesCenter/大客户数据/HW-大连/' + str(year) + '.' + str(month).zfill(2) + '.' + '01' + '-' + str(
        daySub1).zfill(2) + '海王大连.xlsx'  # 新建的excel文件名字
    savefile = xlsxwriter.Workbook(new_excel)  # 新建一个excel表
    new_sheet = savefile.add_worksheet()  # 新建一个sheet表
    for i in range(len(all_data1)):
        for j in range(len(all_data1[i])):
            c = all_data1[i][j]
            new_sheet.write(i, j, c)
    savefile.close()  # 关闭该excel表
    cprint(" > 海王大连数据已导出！ >> 记得发到大连海王流向沟通群>>> ", 'magenta', attrs=['bold', 'reverse', 'blink'])


try:
    DL()  # 大连
except Exception as e:
    print("大连导出出错", e)


def HW():  # 海王
    print('\n' + "开始导出海王数据ing......")
    global sf
    global driver
    global all_data1
    all_data1 = []
    sf = ['北京', '长春', '成都', '大连', '电商', '福州', '广州', '河南', '湖北', '湖南', '深圳总部', '杭州', '江苏', '辽宁', '宁波', '青岛', '潍坊', '上海',
          '沈阳', '深圳', '天津', '泰州']  # 海王省份列表
    wei_zhi = "D:\\FilesCenter\\大客户数据\\HW\\"  # 海王下载路径
    szLocation = 'D:/FilesCenter/大客户数据/HW-深圳/'
    options = webdriver.ChromeOptions()  # 打开设置
    prefs = {'profile.default_content_settings.popups': 0, 'download.default_directory': wei_zhi}  # 设置路径
    options.add_experimental_option('prefs', prefs)
    driver = webdriver.Chrome(executable_path='C:\\Program Files (x86)\\Google\\Chrome\\Application\\chromedriver.exe',
                              options=options)  # 打开浏览器
    driver.implicitly_wait(5)  # 隐式等待
    with open('D://FilesCenter//大客户数据//anquan.txt', 'r', encoding='utf8') as f:
        listCookies = json.loads(f.read())  # 读取cookies
    driver.get('http://srm.nepstar.cn')  # 进入登录界面
    driver.delete_all_cookies()  # 删旧cookies
    for cookie in listCookies:
        if 'expiry' in cookie:
            del cookie['expiry']
        driver.add_cookie(cookie)  # 新增cookies
    time.sleep(3)
    driver.get(
        'http://srm.nepstar.cn/ELSServer_HWXC/default2.jsp?account=110829_1001&loginChage=N&telphone1=18279409642')
    # 读取完cookie刷新页面
    time.sleep(1)
    button_a('//*[@id="treeMenu"]/li[7]/a')  # 数据查询
    time.sleep(0.5)
    button_a('//*[@id="salesOutSourcingInfoManage"]')  # 销售查询
    time.sleep(0.5)
    xf = driver.find_element_by_xpath('/html/body/div[2]/div/nav[2]/div[3]/iframe')  # 先通过xpath定位到iframe
    driver.switch_to.frame(xf)  # 再将定位对象传给switch_to.frame()方法
    button_a('/html/body/div[1]/div[2]/form/div[1]/div/div/span')  # 商品编码
    time.sleep(5)
    driver.switch_to.parent_frame()
    xf = driver.find_element_by_xpath('/html/body/div[4]/div/table/tbody/tr[2]/td/div/iframe')
    driver.switch_to.frame(xf)  # 切换
    button_a('/html/body/div/main/div[1]/div[1]/div[1]/table/thead/tr/th[2]/div/span/input')  # 全选
    time.sleep(0.5)
    button_a('/html/body/div/main/div[2]/button[1]')  # 确定
    time.sleep(0.5)
    driver.switch_to.parent_frame()
    xf = driver.find_element_by_xpath('/html/body/div[2]/div/nav[2]/div[3]/iframe')
    driver.switch_to.frame(xf)  # 切换
    input_b('/html/body/div[1]/div[2]/form/div[2]/div/div/input', time0)  # 查询日期
    time.sleep(0.5)
    input_b('/html/body/div[1]/div[2]/form/div[3]/div/div/input', time1)
    time.sleep(0.5)
    button_a('/html/body/div[1]/div[2]/form/div[4]/div/div/div/p')  # 联采合同
    time.sleep(0.5)
    button_a('/html/body/div[1]/div[2]/form/div[4]/div/div/div/div/ul/li[3]')
    time.sleep(0.5)
    for i in range(2, 24):
        if i == 12:
            cprint("跳过深圳总部", 'cyan', attrs=['bold', 'reverse', 'blink'])
        # elif i == 15:
        #     cprint("无权限访问，跳过辽宁", 'cyan', attrs=['bold', 'reverse', 'blink'])
        else:
            button_a('/html/body/div[1]/div[2]/form/div[5]/div/div/div/p')  # 选地区
            time.sleep(0.5)
            ix = '/html/body/div[1]/div[2]/form/div[5]/div/div/div/div/ul/li[' + str(i) + ']'
            button_a(ix)  # 选好地区
            time.sleep(0.5)
            button_a('/html/body/div[1]/div[2]/form/button[2]')  # 查询
            time.sleep(7)
            button_a('/html/body/div[1]/div[3]/div[1]/div[1]/table/thead/tr/th[2]/div/span')  # 全选
            time.sleep(1)
            button_a('/html/body/div[1]/div[1]/nav/div/div/ul/li[1]/a')  # 查看明细
            time.sleep(1)
            driver.switch_to.parent_frame()  # 回退
            try:
                driver.find_element_by_xpath('/html/body/div[2]/div/nav[1]/ul/li[4]/a')  # 查看明细是否存在
                xf = driver.find_element_by_xpath('/html/body/div[2]/div/nav[2]/div[4]/iframe')  # 切入明细
                driver.switch_to.frame(xf)
                time.sleep(8)
                button_a('/html/body/div/div[1]/nav/div/div/ul/li[2]/a')  # 导出
                driver.switch_to.parent_frame()  # 切出
                time.sleep(6)
                element = driver.find_element_by_xpath('/html/body/div[2]/div/nav[1]/ul/li[4]/a/i[2]')
                ActionChains(driver).move_to_element(element).perform()  # 鼠标悬浮
                button_a('/html/body/div[2]/div/nav[1]/ul/li[4]/a/i[2]')  # 关闭明细
                time.sleep(8)
                all_exce = get_exce(wei_zhi)
                # 得到要合并的所有exce表格数据
                if (all_exce == 0):
                    cprint(sf[i - 2] + "下载出错！！！", 'magenta', attrs=['bold', 'reverse', 'blink'])
                else:
                    for exce in all_exce:
                        fh = xlrd.open_workbook(exce)
                        # 打开文件
                        sheets = get_sheet(fh)
                        # 获取文件下的sheet数量
                        for sheet in range(len(sheets)):
                            row = get_sheetrow_num(sheets[sheet])
                            # 获取一个sheet下的所有的数据的行数
                            all_data1 = get_sheet_data(sheets[sheet], row, i - 2)
                            os.remove(exce)

                print("导出完成！", sf[i - 2])
                xf = driver.find_element_by_xpath('/html/body/div[2]/div/nav[2]/div[3]/iframe')  # 切入查询
                time.sleep(1)
                driver.switch_to.frame(xf)  # 切换
                button_a('/html/body/div[1]/div[3]/div[1]/div[1]/table/thead/tr/th[2]/div/span')  # 全不选

            except:
                cprint(sf[i - 2] + "->列无数据", 'magenta', attrs=['bold', 'reverse', 'blink'])
                xf = driver.find_element_by_xpath('/html/body/div[2]/div/nav[2]/div[3]/iframe')  # 切入查询
                driver.switch_to.frame(xf)  # 切换
                button_a('/html/body/div[1]/div[3]/div[1]/div[1]/table/thead/tr/th[2]/div/span')  # 全不选
            # finally:

    dictCookies = driver.get_cookies()  # 获取cookies
    jsonCookies = json.dumps(dictCookies)
    with open('D://FilesCenter//大客户数据//anquan.txt', 'w') as f:
        f.write(jsonCookies)  # 保存新cookies

    biao_tou = ['类别', '商品SAP编码', '商品名称', '规格', '单位', '店号/区域ID', '店名/区域', '销量', '过账日期', '合同价', '省份']
    all_data1.insert(0, biao_tou)  # 表头写入
    # 下面开始文件数据的写入
    new_excel = "D:\\FilesCenter\\大客户数据\\HW\\" + "HAIWANG" + str(month) + ".xlsx"  # 新建的exce文件名字
    fh1 = xlsxwriter.Workbook(new_excel)  # 新建一个excel表
    new_sheet = fh1.add_worksheet()  # 新建一个sheet表
    for i in range(len(all_data1)):
        for j in range(len(all_data1[i])):
            c = all_data1[i][j]
            new_sheet.write(i, j, c)
    fh1.close()  # 关闭该excel表
    # 导出深圳数据
    hwData = pd.read_excel(wei_zhi + 'HAIWANG' + str(month) + '.xlsx', header=0)
    shenzhenData = hwData[hwData["省份"] == "深圳"]
    if aaa == '1':
        shenzhenData.to_excel(
            szLocation + str(year) + '.' + str(month).zfill(2) + '.' + '01-' + str(day - 2).zfill(2) + "海王深圳.xlsx",
            index=False)
    else:
        shenzhenData.to_excel(
            szLocation + str(year) + '.' + str(month).zfill(2) + '.' + '01-' + str(lastDay) + "海王深圳.xlsx",
            index=False)
    cprint(" > 海王深圳数据已导出！ >> 记得发到深圳海王数据群>>> ", 'magenta', attrs=['bold', 'reverse', 'blink'])


try:
    HW()  # 海王
except Exception as e:
    print("海王导出出错", e)
