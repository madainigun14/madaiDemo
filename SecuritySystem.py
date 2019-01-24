# -*- coding: utf-8 -*-
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import NoAlertPresentException
import unittest, time, re
import Util

username = "baomi"
password = "123456"

def login(driver):
    driver.get("http://192.168.0.22:8801/sec/a/secuserManage/userMain?menuId=79f564de2dcc49dba80ba713a87a0913")
    try:
        driver.switch_to.alert.accept()
    except Exception as e:
        pass
    driver.find_element_by_id("username").clear()
    driver.find_element_by_id("password").clear()
    driver.find_element_by_id("username").send_keys(username)
    driver.find_element_by_id("password").send_keys(password)
    driver.find_element_by_css_selector("fieldset>div[class='clearfix']>button").click()
    driver.implicitly_wait(20)


class SecuritySystem(unittest.TestCase):
    def setUp(self):
        self.driver = webdriver.Chrome()
        self.driver.maximize_window()
        self.driver.implicitly_wait(10)
        self.base_url = "http://192.168.0.22:8801/opt/a/cas"
        self.verificationErrors = []
        self.accept_next_alert = True

    def addExistingPosition(self):
        """
        添加一个岗位，这个岗位已经在岗位管理里存在
        :return:
        """
        driver = self.driver
        login(driver)
        time.sleep(1)

        # 涉密岗位信息 F:\涉密信息管理系统\Test Data\涉密岗位信息.xlsx
        # 初始化Excel对象，用来接收传过来的Json
        excelobj = Util.Excel("F:\涉密信息管理系统\Test Data\涉密岗位信息.xlsx")
        sheetdata = excelobj.readsheet("涉密岗位信息")
        print("+++++++++++++++++++sheetdata++++++++++++++++++")
        print(sheetdata)
        print("+++++++++++++++++++sheetdata++++++++++++++++++")
        print("+++++++++++++++++++部门信息++++++++++++++++++")
        for i in range(0, len(sheetdata)-1):
            print(sheetdata[i]["单位部门"])
        print("+++++++++++++++++++部门信息++++++++++++++++++")

        # 在左侧树结构菜单中找到“岗位管理”所在的index（css选择器的索引），然后点击
        index = str(Util.BootStrap(driver)._getcssselectorindex("sidebar", "岗位管理"))
        driver.find_element_by_css_selector("div#sidebar>ul>li:nth-child(" + index + ")").click()
        time.sleep(1)

        # 点击“新建”
        driver.find_element_by_css_selector("form#postSearchForm>div>input:nth-child(1)").click()
        time.sleep(1)

        # 添加“岗位名称”
        driver.find_element_by_id("postName").clear()
        driver.find_element_by_id("postName").send_keys(u"成都市公安局保密专员")

        # 选择“涉密等级”
        Select(driver.find_element_by_id("secretGrade")).select_by_visible_text(u"核心涉密")

        # 点击“添加”按钮，添加“确定依据”
        driver.find_element_by_xpath("//button[@onclick='addGistRow()']").click()
        driver.find_element_by_css_selector("input.input.width-100").clear()
        driver.find_element_by_css_selector("input.input.width-100").send_keys(u"保密法")

        # 选择单位部门
        driver.find_element_by_id("chooseUnitStr").click()
        time.sleep(1)

        nodeid = driver.find_element_by_css_selector("a[title='发改委']").get_property("id")
        nodeid = nodeid.replace("a", "switch")
        driver.find_element_by_id(nodeid).click()
        time.sleep(1)

        subnodeid = driver.find_element_by_css_selector("a[title='发改委南京分舵']").get_property("id")
        subnodeid = subnodeid.replace("a", "check")
        driver.find_element_by_id(subnodeid).click()
        time.sleep(1)

        # 点击“确定”，保存选择的树节点
        driver.find_element_by_xpath("//button[@onclick='saveChooseUnit()']").click()
        time.sleep(2)

        # 点击“确定”提交整个页面表单
        driver.find_element_by_xpath("//button[@onclick='saveSecPost()']").click()
        time.sleep(2)

        # 如果添加成功，则会在table里置顶处添加一条数据，判断的css选择器： table#sample-table-1>tbody>tr>td:nth-child(2)
        # 如果添加失败，会弹出层，css选择器是div[class='modal-dialog modal-sm']>div>div:nth-child(2)>div
        # 文本是  "保存失败.岗位名称已存在"
        is_displayed = driver.find_element_by_css_selector("div[class='modal-dialog modal-sm']>div>div:nth-child(2)>div").is_displayed()
        is_textmatch = driver.find_element_by_css_selector("div[class='modal-dialog modal-sm']>div>div:nth-child(2)>div").text == "保存失败.岗位名称已存在"
        self.assertEqual(is_displayed and is_textmatch, True, "添加失败的div未弹出来，出现异常了，用例失败！")

    def addPosition(self):
        """
        添加一个岗位，在“岗位管理”里不存在
        :return:
        """
        driver = self.driver
        login(driver)
        time.sleep(1)

        # 涉密岗位信息 F:\涉密信息管理系统\Test Data\涉密岗位信息.xlsx
        # 初始化Excel对象，用来接收传过来的Json
        excelobj = Util.Excel("F:\涉密信息管理系统\Test Data\涉密岗位信息.xlsx")
        sheetdata = excelobj.readsheet("涉密岗位信息")
        print("+++++++++++++++++++sheetdata++++++++++++++++++")
        print(sheetdata)
        print("+++++++++++++++++++sheetdata++++++++++++++++++")
        print("+++++++++++++++++++部门信息++++++++++++++++++")
        for i in range(0, len(sheetdata)-1):
            print(sheetdata[i]["单位部门"])
        print("+++++++++++++++++++部门信息++++++++++++++++++")

        Util.BootStrap(driver).gotomenu("岗位管理")

        # 点击“新建”
        driver.find_element_by_css_selector("form#postSearchForm>div>input:nth-child(1)").click()
        time.sleep(1)

        # 添加“岗位名称”
        driver.find_element_by_id("postName").clear()
        driver.find_element_by_id("postName").send_keys(u"成都市公安局保密专员")

        # 选择“涉密等级”
        Select(driver.find_element_by_id("secretGrade")).select_by_visible_text(u"核心涉密")

        # 点击“添加”按钮，添加“确定依据”
        driver.find_element_by_xpath("//button[@onclick='addGistRow()']").click()
        driver.find_element_by_css_selector("input.input.width-100").clear()
        driver.find_element_by_css_selector("input.input.width-100").send_keys(u"保密法")

        # 选择单位部门
        driver.find_element_by_id("chooseUnitStr").click()
        time.sleep(1)

        nodeid = driver.find_element_by_css_selector("a[title='发改委']").get_property("id")
        nodeid = nodeid.replace("a", "switch")
        driver.find_element_by_id(nodeid).click()
        time.sleep(1)

        subnodeid = driver.find_element_by_css_selector("a[title='发改委南京分舵']").get_property("id")
        subnodeid = subnodeid.replace("a", "check")
        driver.find_element_by_id(subnodeid).click()
        time.sleep(1)

        # 点击“确定”，保存选择的树节点
        driver.find_element_by_xpath("//button[@onclick='saveChooseUnit()']").click()
        time.sleep(2)

        # 点击“确定”提交整个页面表单
        driver.find_element_by_xpath("//button[@onclick='saveSecPost()']").click()
        time.sleep(2)

        # 如果添加成功，则会在table里置顶处添加一条数据，判断的css选择器： table#sample-table-1>tbody>tr>td:nth-child(2)
        # 如果添加失败，会弹出层，css选择器是div[class='modal-dialog modal-sm']>div>div:nth-child(2)>div
        # 文本是  "保存失败.岗位名称已存在"
        is_displayed = driver.find_element_by_css_selector("div[class='modal-dialog modal-sm']>div>div:nth-child(2)>div").is_displayed()
        is_textmatch = driver.find_element_by_css_selector("div[class='modal-dialog modal-sm']>div>div:nth-child(2)>div").text == "保存失败.岗位名称已存在"
        self.assertEqual(is_displayed and is_textmatch, True, "添加失败的div弹出来了")

    def tearDown(self):
        self.driver.quit()
        self.assertEqual([], self.verificationErrors)


if __name__ == "__main__":
    suite = unittest.TestSuite()
    suite.addTest(SecuritySystem("addPosition"))

    runner = unittest.TextTestRunner(failfast=False)
    runner.run(suite)
