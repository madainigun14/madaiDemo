# -*- coding: utf-8 -*-
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import NoAlertPresentException
import unittest
import time
import json
import Util
import random
import traceback


url_1 = "http://192.168.0.22:8801/exam/a/paper/etExamPaper?menuId=8f6897807d7840898f729bf2b6fa9401&title=%E5%9F%B9%E8%AE%AD%E7%AE%A1%E7%90%86"
url_2 = "http://192.168.0.22:8801/exam/a/course/etCourseInfo?menuId=0c3f2b815ab24f1d8d9a874fd41ebb41&title=%E8%AF%BE%E7%A8%8B%E7%AE%A1%E7%90%86"
url_3 = "http://192.168.0.22:8801/sec/a/secuserManage/userMain?menuId=79f564de2dcc49dba80ba713a87a0913"
url_news = "http://192.168.0.22:8801/opt/a/sys/log/platform?menuId=c4cbcde27ccf40b9af8e7a2c108dd534"

# 登录平台
def login(driver):
    driver.find_element_by_id("username").clear()
    driver.find_element_by_id("password").clear()
    driver.find_element_by_id("username").send_keys("baomi")
    driver.find_element_by_id("password").send_keys("123456")
    driver.find_element_by_css_selector("fieldset>div[class='clearfix']>button").click()
    time.sleep(1)


class ZTreeTest(unittest.TestCase):
    def setUp(self):
        self.driver = webdriver.Chrome()
        self.driver.maximize_window()
        self.driver.implicitly_wait(30)
        self.verificationErrors = []
        self.accept_next_alert = True

    def addpaper(self):
        driver = self.driver
        driver.get(url_1)

        login(driver)  # 登录

        try:
            bootstrap = Util.BootStrap(driver)
            bootstrap.goto("添加试卷")

            ztree = Util.ZTree(driver)

            ztree.checknodebyvalue("培训1")
            ztree.checknodebyvalue("保密常识")
            ztree.checknodebyvalue("常识1")
            ztree.checknodebyvalue("保密培训")
            ztree.unchecknodebyvalue("保密培训")
            time.sleep(3)
            ztree.deselect()

            print(ztree._getvalues(Util.TreeNodeLevel.LEVEL1))
            print(ztree._getvalues(Util.TreeNodeLevel.LEVEL2))
            print(ztree._getvalues(Util.TreeNodeLevel.ALL))

            time.sleep(1)
            ztree.expandall()
            ztree.collideall()
            ztree.expandbyindex(2)
            ztree.collidebyindex(2)
            time.sleep(3)

            print("开始执行javascript")
            js0 = 'alert(10086)'
            js1 = "alert(document.getElementsByClassName('modal-body')[0].scrollHeight)"
            js2 = "alert(document.getElementsByClassName('modal-body')[0].scrollWidth)"
            js3 = "var q=document.getElementsByClassName('modal-body')[0].scrollTop=200"
            # driver.execute_script(js0)
            # driver.execute_script(js1)
            # driver.execute_script(js2)
            driver.execute_script(js3)

            time.sleep(10)

        except Exception as e:
            traceback.print_exc()
            print("◆◆◆◆出现异常了◆◆◆◆ \r\n :" + str(e))

    def selecttest(self):
        driver = self.driver
        driver.get(self.base_url)

        login(driver)  # 登录

        optionvalue = "审查不通过"

        bootstrap = Util.BootStrap(driver)
        bootstrap.selectbytext(optionvalue)

        time.sleep(10)

    def log_pagination(self):
        driver = self.driver
        driver.get(url_news)

        bootstrap = Util.BootStrap(driver)
        bootstrap.login(url_news, "shenji", "123456")  # 登录

        js = "onclick=page(100,10,'');"
        driver.execute_script(js)
        time.sleep(2)

        js1 = "onclick=page(1,10,'');"
        driver.execute_script(js1)
        time.sleep(2)

        paging = Util.Paging(driver)
        print(paging.getmaxpage())

    def tearDown(self):
        self.driver.quit()
        self.assertEqual([], self.verificationErrors)


if __name__ == "__main__":
    suite = unittest.TestSuite()
    # suite.addTest(ZTreeTest("addpaper"))
    suite.addTest(ZTreeTest("log_pagination"))

    runner = unittest.TextTestRunner(failfast=False)
    runner.run(suite)