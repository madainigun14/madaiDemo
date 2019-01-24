# -*- coding: utf-8 -*-
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import NoAlertPresentException
import unittest, time, re
import Util

url = "http://192.168.0.22:8801/opt/a/sec/org?menuId=53c6606a147246499a2efe841b8dd3b9"
username = "system"
password = "123456"

class SystemManage(unittest.TestCase):
    def setUp(self):
        self.driver = webdriver.Chrome()
        self.driver.maximize_window()
        self.driver.implicitly_wait(10)
        self.base_url = "http://192.168.0.22:8801/opt/a/cas"
        self.verificationErrors = []
        self.accept_next_alert = True

    # 添加机构用户
    def addOrgUser(self):
        driver = self.driver
        bootstrap = Util.BootStrap()
        bootstrap.login(driver, url, username, password)

        # 涉密岗位信息 F:\涉密信息管理系统\Test Data\涉密岗位信息.xlsx
        # 初始化Excel对象，用来接收传过来的Json
        excelobj = Util.Excel("F:\中纪委综合业务管理平台\Test Data\系统管理.xlsx")
        sheetdata = excelobj.readsheet("机构用户")
        for i in range(0, len(sheetdata)):
            time.sleep(0.5)

            # 点击“新建”
            driver.find_element_by_css_selector("div.col-xs-3>button:nth-child(1)").click()
            time.sleep(1)

            # 用户姓名
            driver.find_element_by_id("name").send_keys(sheetdata[i]["用户姓名"])

            # 用户账号
            driver.find_element_by_id("loginName").send_keys(sheetdata[i]["用户账号"])

            #联系方式
            driver.find_element_by_id("mobile").send_keys(sheetdata[i]["联系方式"])

            #身份证号
            driver.find_element_by_id("idCard").send_keys(sheetdata[i]["身份证号"])

            #邮箱
            driver.find_element_by_id("email").send_keys(sheetdata[i]["邮箱账号"])

            #备注
            driver.find_element_by_css_selector("textarea[name='remarks']").send_keys(sheetdata[i]["备注"])

            #点击“确定”提交表单
            driver.find_element_by_css_selector("button[title='确定']").click()

            time.sleep(0.5)

            for j in range(0, len(driver.find_elements_by_css_selector("[class='required']"))):
                txtvalue = driver.find_elements_by_css_selector("[class='required']")[j].text
                print(txtvalue)

        # 2级部门的唯一查找方式  ul[id='tree-unit_1_ul']>li>a[title='国土资源局']
        # 3级部门的唯一查找方式  li[id='tree-unit_1']>ul>li>ul>li>a[title='呼和浩特市公安局']
        # 然而，并没有标识某个单位是2级还是3级，所以需要换一种定位方式a[title='']


    def tearDown(self):
        self.driver.quit()
        self.assertEqual([], self.verificationErrors)


if __name__ == "__main__":
    suite = unittest.TestSuite()
    suite.addTest(SystemManage("addOrgUser"))

    runner = unittest.TextTestRunner(failfast=False)
    runner.run(suite)
