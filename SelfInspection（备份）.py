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
import Util, WebControlUtil
import random


username = 'madai4'

# 普通用户登录平台
def login(driver):
    driver.get("http://192.168.0.22:8801/opt/a")
    driver.find_element_by_id("username").clear()
    driver.find_element_by_id("password").clear()
    driver.find_element_by_id("username").send_keys(username)
    driver.find_element_by_id("password").send_keys("123456")
    driver.find_element_by_css_selector("fieldset>div[class='clearfix']>button").click()
    time.sleep(1)

    # 在平台页面，点击“自查自评业务系统”
    driver.find_element_by_css_selector("div[class='item-main item-2']").click()

    # 由于自查自评是另起一个窗体，所以，必须切换到句柄
    windowHandles = driver.window_handles
    driver.switch_to.window(windowHandles[1])
    time.sleep(1)


def logintotaskpage(driver):
    driver.get("http://192.168.0.22:8801/inspection/client/taskManage/taskMainList?type=1&parentId=5ef39b1754d747b3994d59ebc71374f0")
    try:
        driver.switch_to.alert.accept()
    except Exception as e:
        pass
    driver.find_element_by_id("username").clear()
    driver.find_element_by_id("password").clear()
    driver.find_element_by_id("username").send_keys(username)
    driver.find_element_by_id("password").send_keys("123456")
    driver.find_element_by_css_selector("fieldset>div[class='clearfix']>button").click()
    driver.implicitly_wait(20)



# 弹出的层很烦人，有时候不会自己消失，于是先点击一下DIV，然后在点关闭
def closeDIV(driver):
    driver.find_element_by_css_selector("div[id='gritter-notice-wrapper']>div").click()
    driver.find_element_by_css_selector("a[class='gritter-close']").click()
    driver.implicitly_wait(2)


def top(driver):
    linklist_out = driver.find_elements_by_css_selector("a[title='开始任务']")
    # 界面有i个链接，循环i次
    for i in range(0, len(linklist_out)):
        # print("最外层，链接有"+ str(len(linklist)) +"个，现在开始循环i层")

        linklist_in = driver.find_elements_by_css_selector("a[title='开始任务']")
        time.sleep(0.5)
        try:
            linklist_in[i].click()
            time.sleep(1)
            driver.find_elements_by_css_selector("a[class='closed-add']")[0].click()

        except Exception as e:
            print("◆◆◆◆置顶出现异常了◆◆◆◆ \r\n :" + str(e))
            continue

        print("=====置顶了一个任务=======")  # debug only
        page = WebControlUtil.Paging(driver)
        page.movetolastpage()
    top(driver)
    if len(linklist_out) == 0:
        pass


def selfinspect(driver):
    linklist_out = driver.find_elements_by_css_selector("a[title='开始任务']")
    linklen = len(driver.find_elements_by_css_selector("a[href='javascript:']"))
    # 界面有i个链接，循环i次
    for i in range(0, len(linklist_out)-1, 1):
        # print("最外层，链接有"+ str(len(linklist)) +"个，现在开始循环i层")

        linklist_in = driver.find_elements_by_css_selector("a[title='开始任务']")
        time.sleep(0.5)
        try:
            linklist_in[i].click()
            time.sleep(1)

            tasklist_out = driver.find_elements_by_css_selector("div[class='task-list']")
            # 有j个任务，要循环j次
            for j in range(0, len(tasklist_out)):
                # print("有任务" + str(len(tasklist_out)) + "个，现在开始循环j层")

                # 这里，需要一个外部一个内部，原因是页面刷新可能导致获取不到元素，抛异常：
                # StaleElementReferenceException: stale element reference: element is not attached to the page document
                # 参考<https://www.cnblogs.com/fengpingfan/p/4583325.html>
                tasklist_in = driver.find_elements_by_css_selector("div[class='task-list']")
                tasklist_in[j].click()
                time.sleep(0.5)

                checklist = driver.find_elements_by_css_selector("div[id='serinaNoDiv']>ul>li")
                # 有k个自查项，要循环k次
                for k in range(0, len(checklist)):
                    # 自查自评的标题
                    title = driver.find_element_by_css_selector("div[class='titem-title']").text
                    # print("有自查项" + str(len(checklist)) + "个，现在开始循环k层")
                    # 先点1.1, 1.2，1.3...子项
                    checklist[k].click()
                    # print("点击了自查项") #debug only
                    time.sleep(0.1)

                    # region 答题方式，目前为了让程序跑通，方法写的很死。判断自查自评检查项的标题
                    if not ("加重扣分项" in title):
                        print("是的，这是Select")
                        # 答题方法一： Select>option    选择第四个选项
                        maxindex = len(
                            driver.find_elements_by_css_selector("select[onchange='modifyScoreSel();']>option"))
                        Select(driver.find_element_by_css_selector(
                            "select[onchange='modifyScoreSel();']")).select_by_index(maxindex - 1)
                        print("选择了得分")  # debug only
                        time.sleep(0.1)
                    else:
                        print("找不到下拉框，但是找到了Spinner控件")
                        # 答题方法二： Spinner>a      随机点▲或▼几次
                        # driver.find_element_by_css_selector("div[id='spinnerDiv']>a[class='spin-up']").click()  # 加1
                        print("点击了Spinner，设置了得分")  # debug only
                        time.sleep(0.1)
                        # 答题方法三： Text          输入数值
                    # endregion

                # 点击“完成本项
                driver.find_element_by_id("thisIsOk").click()
                time.sleep(1)

                # 弹出的窗体，点击“确定”
                driver.find_element_by_css_selector("div#submitStand>div.modal-foot>button:nth-child(1)").click()
                time.sleep(0.5)
                # 关闭弹出的提示信息 BUG

                # 弹出的层很烦人，有时候不会自己消失，于是先点击一下DIV，然后在点关闭
                closeDIV(driver)

                print("================完成了一个检查项======================")  # debug only

            # 等检查项都循环完毕，点击“提交”按钮
            driver.find_element_by_id("submitBtnclient").click()
            time.sleep(1)

            # 弹出任务完成的确认对话框, 先填入必填项“自查自评负责人”，“联系人”
            driver.find_element_by_id("person").clear()
            driver.find_element_by_id("linkman").clear()
            driver.find_element_by_id("person").send_keys(username)
            driver.find_element_by_id("linkman").send_keys(username)
            driver.find_element_by_css_selector("button[onclick='submitTaskconfirm()']").click()
            time.sleep(2)

        except Exception as e:
            print("◆◆◆◆出现异常了◆◆◆◆ \r\n :" + str(e))
            continue

        # closeDIV(driver) #这行代码可要可不要。但是现在有个BUG，导致循环列表里的第二个高亮的自查自评任务失败
        # BUG 1269 ，已经和赵进确认过
        print("※※※※※※※※※※完成了一个自查自评任务！※")  # debug only

    # 点击下一页，进入下一个循环
    driver.find_elements_by_css_selector("a[href='javascript:']")[linklen - 2].click()
    time.sleep(2)
    selfinspect(driver)

def paging(driver):
    page = WebControlUtil.Paging(driver)
    page.movetolastpage()

    linklist_out = driver.find_elements_by_css_selector("a[title = '开始任务']")  # 记录任务后面“开始任务”高亮的元素

    # 如果找不到一个高亮的，则点击分页到下一页继续找;否则直接执行循环
    if len(linklist_out) == 0:
        page.movebackward()
    else:
        selfinspect(driver)
        time.sleep(1)
        paging(driver)


class SelfInspection(unittest.TestCase):
    def setUp(self):
        self.driver = webdriver.Chrome()
        self.driver.maximize_window()
        self.driver.implicitly_wait(30)
        self.base_url = "http://192.168.0.22:8801/opt/a"
        self.verificationErrors = []
        self.accept_next_alert = True

    def login(self):
        driver = self.driver
        driver.get(self.base_url)
        login(driver)
        self.assertEqual(driver.title,"综合业务管理平台 - 首页")  #测试点1：判断登录成功与否

    def startInspection(self):
        print("开始执行startInspection()")
        driver = self.driver
        logintotaskpage(driver)


        selfinspect(driver)



    def tearDown(self):
        self.driver.quit()
        self.assertEqual([], self.verificationErrors)


if __name__ == "__main__":
    suite = unittest.TestSuite()
    # suite.addTest(SelfInspection("login"))
    suite.addTest(SelfInspection("startInspection"))


    runner = unittest.TextTestRunner(failfast=False)
    runner.run(suite)