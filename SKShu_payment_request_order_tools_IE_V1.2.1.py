import easygui
import os
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By                             # 定位类型八种定位方式
import pyperclip as pc
import time
from selenium.webdriver.support.select import Select

IEDriver = "C:\Program Files\Internet Explorer\IEDriverServer.exe"
edgeDriver = "C:\Program Files\Python38\msedgedriver.exe"
# webdriver = webdriver.Ie(IEDriver)
webdriver = webdriver.Ie()
global plantList
plantList = ["三棵树涂料股份", "福建三棵树建筑材料", "四川三棵树涂料", "上海三棵树防水技术", "福建三江包装", "天津三棵树涂料"
    , "河南三棵树涂料", "福建三棵树建筑装饰有限公司", "福建三棵树物流有限公司", "安徽三棵树涂料有限公司", "河北三棵树涂料有限公司"
    , "莆田市禾三投资有限公司", "莆田三棵树涂料贸易有限公司", "福建省三棵树新材料有限公司", "上海三棵树新材料科技有限公司", "上海春之葆生物科技有限公司"
    , "小森新材料科技有限公司", "三棵树（上海）建筑材料有限公司", "三棵树（上海）建筑材料有限公司", "莆田市三棵树公益基金会", "莆田市金淼慈善基金会", "三棵树（上海）新材料研究有限公司"
    , "广州大禹防漏技术开发有限公司", "大禹九鼎新材料科技有限公司", "广州大禹九鼎新材料有限公司", "湖北大禹九鼎新材料科技有限公司", "北京三棵树新材料科技有限公司", "三棵树集团国际有限公司"
    , "湖北三棵树新材料科技有限公司", "河南三棵树新材料科技有限公司", "江苏麦格美节能科技有限公司", "廊坊富达新型建材有限公司", "3TREES PAINT SDN. BHD."
    , "福建三棵树电子商务有限公司", "福建三棵树电子商务有限公司莆田隔热材料销售分公司", "福建三棵树电子商务有限公司莆田建筑材料销售分公司", "福建三棵树电子商务有限公司莆田防水材料销售分公司"
    , "福建三棵树电子商务有限公司莆田隔音材料销售分公司", "福建三棵树电子商务有限公司莆田互联网销售分公司", "福建三棵树电子商务有限公司莆田涂料销售分公司", "福建三棵树电子商务有限公司厦门分公司",
             "福建三棵树教育科技有限公司", "", "", ""]


def acc():
    # open website
    url = 'http://purch.skshu.com.cn/index.jsp'
    webdriver.get(url)
    webdriver.maximize_window()
    windows1 = webdriver.window_handles
    """Login"""
    pc.copy("YL000658")
    webdriver.find_element_by_xpath('//*[@id="username"]').send_keys(Keys.CONTROL, 'v')
    pc.copy("Basf2020")
    webdriver.find_element_by_xpath('//*[@id="password"]').send_keys(Keys.CONTROL, 'v')
    webdriver.find_element_by_xpath('//*[@id="submitbutton_10904_1"]').click()
    """供应商请款单（A类）"""
    webdriver.find_element_by_xpath('//*[@id="leftMenuTree_7_span"]').click()
    time.sleep(2)
    #   Supplier payment request
    # webdriver.find_element_by_id('leftMenuTree_12_span').click()
    webdriver.find_element_by_id('leftMenuTree_7_ul')
    webdriver.find_element_by_id('leftMenuTree_12').click()
    webdriver.find_element_by_xpath('//*[@id="leftMenuTree_12_span"]').click()
    time.sleep(10)
    """Ending """

    """Create new payment request order """
    time.sleep(5)
    easygui.ynbox("请手动点击新建供应商请款单", "Textbox")
    time.sleep(5)
    window_handles = webdriver.window_handles
    webdriver.switch_to.window(window_handles[-1])
    #webdriver.implicitly_wait(5)


    "switch to he latest window"
    # webdriver.find_element_by_xpath('//*[@id="field_8a8ad0944c2a49d5014c2c06771e15f2"]').click()
    order_plant = easygui.choicebox("请选择要新建请款单的工厂", "请款单", plantList)

    print(order_plant)

    """工厂选择-- 浏览器原因未能实现"""
    webdriver.find_element_by_id('guide-step')
    webdriver.find_element_by_tag_name('table')
    webdriver.find_element_by_tag_name('tbody')
    webdriver.find_element_by_tag_name('tr')
    webdriver.find_element_by_id('rightTD')
    webdriver.find_element_by_id('rightDiv')
    webdriver.find_element_by_id('content')
    framelist = webdriver.find_elements_by_tag_name('iframe')
    for f in [0,len(framelist)]:
        try:
            webdriver.switch_to.frame(f)
            if webdriver.find_element_by_xpath('/html/body/form'):
                print('succeuss')
        except:
            webdriver.switch_to.parent_frame()



    window_handles = webdriver.window_handles
    webdriver.switch_to.window(window_handles[-1])

    webdriver.find_element_by_xpath('/html')
    # webdriver.find_element_by_id('ext-gen8')

    #  webdriver.find_element_by_xpath('/html/body')
    # plant_locater = (By.ID, 'EweaverForm')
    # WebDriverWait(webdriver, 10, 1).until(EC.presence_of_element_located((plant_locater)))
    # webdriver.find_element_by_xpath('//*[@id="EweaverForm"]')  #######


    formArea = webdriver.execute_script('document.getElementById("EweaverForm")')

    # btnPlant = Select(webdriver.find_element_by_class_name("InputStyle6"))

    webdriver.execute_script('document.getElementById("tabPanel")')
    webdriver.execute_script('document.getElementById("ext-gen15")')
    webdriver.execute_script('document.getElementById("ext-gen16")')
    webdriver.execute_script('document.getElementById("")')

    # webdriver.find_element_by_xpath('//*[@id="EweaverForm"]')
    # webdriver.find_element_by_id(r'EweaverForm')
    # webdriver.find_element_by_css_selector('action')    #无法定位该元素

    formArea.find_element_by_id('tabPanel')
    webdriver.find_element_by_id('ext-gen15')

    webdriver.find_element_by_id('ext-gen16')
    webdriver.find_element_by_id('ext-comp-1004')
    webdriver.find_element_by_id('ext-comp-1003')
    time.sleep(0.5)
    webdriver.find_element_by_id('ext-gen28')
    webdriver.find_element_by_id('ext-gen30')
    webdriver.find_element_by_id('tab1')
    time.sleep(0.5)
    webdriver.find_element_by_id('layoutFrame')
    webdriver.find_element_by_tag_name('center')
    webdriver.find_element_by_id('layoutDiv')
    webdriver.find_element_by_id('ext-gen71')
    webdriver.find_element_by_tag_name('tbody')
    webdriver.find_element_by_tag_name('tr')
    webdriver.find_element_by_id('ext-gen64')
    selectPlant = Select(webdriver.find_element_by_class_name('InputStyle6'))

    selectPlant.select_by_value(order_plant)


if __name__ == '__main__':
    acc()
