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

# 登录平台
def login(driver):
    driver.find_element_by_id("username").clear()
    driver.find_element_by_id("password").clear()
    driver.find_element_by_id("username").send_keys("baomi")
    driver.find_element_by_id("password").send_keys("123456")
    driver.find_element_by_css_selector("fieldset>div[class='clearfix']>button").click()
    time.sleep(1)


class ExamManager(unittest.TestCase):
    def setUp(self):
        self.driver = webdriver.Chrome()
        self.driver.maximize_window()
        self.driver.implicitly_wait(30)
        self.base_url = "http://192.168.0.22:8801/opt/a/cas"
        self.verificationErrors = []
        self.accept_next_alert = True

    def addpaper(self):
        driver = self.driver
        driver.get("http://192.168.0.22:8801/exam/a/paper/etExamPaper?menuId=8f6897807d7840898f729bf2b6fa9401&title=%E5%9F%B9%E8%AE%AD%E7%AE%A1%E7%90%86")

        login(driver)  # 登录

        try:
            bootstrap = Util.BootStrap(driver)
            bootstrap.goto("添加试卷")
            bootstrap.goto("培训1")
            time.sleep(10)

        except Exception as e:
            traceback.print_exc()
            print("◆◆◆◆出现异常了◆◆◆◆ \r\n :" + str(e))

    def addExam(self):
        print("开始执行addExam()")
        driver = self.driver
        driver.get("http://192.168.0.22:8801/exam/a/exam/etExam?menuId=92&title=%E8%80%83%E8%AF%95%E7%AE%A1%E7%90%86")

        login(driver)        # 登录

        examname = "考试_自动化"
        papername = "English 卷 I"
        datetime_start = "2025/2/10"
        datetime_end = "2025/3/1"
        try:
            # 点击发起考核
            # driver.find_element_by_css_selector("button[onclick='showEditExam()']").click()
            bootstrap = Util.BootStrap(driver)
            bootstrap.goto("考试管理")
            bootstrap.goto("发起考核")
            time.sleep(1)

            # 考试名称
            driver.find_element_by_id("examTitle").send_keys(examname)

            # 选择考试类型
            bootstrap.goto("移动平板考试")

            # 选择试卷
            index = Util.BootStrap(driver).getselectindex(papername)
            driver.find_element_by_id("paperId").find_element_by_xpath()

            # ####开始时间
            driver.find_element_by_id("startDate").click()
            driver.implicitly_wait(5)
            iframe_startdate = driver.find_element_by_tag_name("iframe")
            driver.switch_to.frame(iframe_startdate)    # 切到iframe
            driver.implicitly_wait(5)

            Util.DateTime(datetime_start).selectdatetime(driver)

            driver.switch_to.default_content()      # 切回
            time.sleep(1)

            # ####结束时间
            driver.find_element_by_id("endDate").click()
            driver.implicitly_wait(5)
            iframe_enddate = driver.find_element_by_tag_name("iframe")
            driver.switch_to.frame(iframe_enddate)  # 切到iframe
            driver.implicitly_wait(5)

            # 调用Util里的方法，进行选择
            Util.DateTime(datetime_end).selectdatetime(driver)

            driver.switch_to.default_content()      # 切回
            driver.implicitly_wait(5)

            # 点击“确定”，提交表单
            driver.find_element_by_css_selector("input[class='btn btn-xs btn-danger'][title='确定']").click()
            time.sleep(2)

            result = driver.find_element_by_css_selector("div[class='bootbox-body']").text
            self.assertEqual("保存信息成功", result)
            # driver.find_element_by_xpath("/html/body/div[10]/div/div/div[3]/button").click()
            # driver.find_element_by_id("examTitle").send_keys(examname+"0001")

        except Exception as e:
            traceback.print_exc()
            print("◆◆◆◆出现异常了◆◆◆◆ \r\n :" + str(e))

    def updateexam(self):
        print("开始执行previewNews()")
        driver = self.driver
        driver.get(self.base_url)
        login(driver)

        # 在平台页面，点击“新闻管理”
        driver.find_element_by_css_selector("div[class='item-main item-5']").click()
        time.sleep(1)

        # 初始化Excel对象，用来接收传过来的Json
        excelobj = Util.Excel("F:\新闻管理系统\Test Data\新闻管理.xlsx")
        sheetdata = excelobj.readSheet("添加新闻")

        news_dict = dict()
        for i in range(0, len(sheetdata)):
            news_dict[sheetdata[i]["新闻标题"].strip()] = sheetdata[i]["正文"].strip()

        # print(news_dict.__str__())

        tr_list = driver.find_elements_by_css_selector("tbody>tr")
        for i in range(0, len(tr_list)):
            # print("第"+str(i)+"个tr:")
            # print(tr_list[i].text)
            td_list = driver.find_elements_by_css_selector("tbody>tr")[i].find_elements_by_tag_name("td")

            # 循环遍历每个tr里的td
            # for j in range(0, len(td_list)):
            #     print("第" + str(j) + "个td的值是：" + td_list[j].text)

            # 获取第2个td“新闻标题”，根据sheetdata查找出相应的“正文”
            td_news_title = td_list[2].text
            # print("新闻标题 "+ str(i) + " 是：：：" + td_news_title)
            # print("字典的Value，“正文”是：" + news_dict[td_list[2].text])

            # 获取第8个td，也就是最后面的4个操作按钮

            # 点击“详情”按钮，进行preview
            try:
                driver.find_elements_by_css_selector("tbody>tr>td:nth-child(9)>div>a[title='详情']")[i].click()
                print("点开了一则新闻，标题是：" + str(td_news_title))
                time.sleep(1)
                news_title_preview = driver.find_element_by_tag_name("h3").text
                self.assertEqual(news_title_preview, td_list[2].text, "预览标题与之前插入的标题不符合，预览功能有BUG！")
                time.sleep(3)
                driver.find_element_by_id("view-modal").find_element_by_css_selector("a.closed-add").click()
                time.sleep(1)
            except Exception as e:
                print("◆◆◆◆出现异常了◆◆◆◆ \r\n :" + str(e))
                continue
        time.sleep(10)

    def auditexam(self):
        print("开始执行auditNews()")
        driver = self.driver
        driver.get(self.base_url)
        login(driver)

        # 在平台页面，点击“新闻管理”
        driver.find_element_by_css_selector("div[class='item-main item-5']").click()
        time.sleep(1)

        # 初始化Excel对象，用来接收传过来的Json
        excelobj = Util.Excel("F:\新闻管理系统\Test Data\新闻管理.xlsx")
        sheetdata = excelobj.readSheet("添加新闻")

        news_dict = dict()
        for i in range(0, len(sheetdata)):
            news_dict[sheetdata[i]["新闻标题"].strip()] = sheetdata[i]["正文"].strip()

        # print(news_dict.__str__())

        for i in range(0, len(driver.find_elements_by_css_selector("tbody>tr"))):

            try:
                driver.find_elements_by_css_selector("tbody>tr>td:nth-child(9)>div>a[title='审核']")[i].click()
                print("点开了审核按钮")
                time.sleep(2)

                # value=1, 审核通过  value=2 审核不通过
                randvalue = random.randint(1, 1)
                driver.find_elements_by_css_selector("div.list-index>div.list-radio>label")[randvalue].click()
                time.sleep(1)

                self.assertEqual(driver.find_element_by_css_selector("div#audit-modal-content>div>h4").text, "审核状态", "新闻审核窗口弹出失败，有BUG！")

                driver.find_element_by_css_selector("button[onclick='saveStatus()']").click()
                time.sleep(1)

                # 弹出“操作完成”提示，点确定
                driver.find_element_by_xpath("/html/body/div[7]/div/div/div[3]/button").click()
                time.sleep(1)

                # 如何审核通过，则状态改成“审核通过”，反之“审核不通过”
                td_list = driver.find_elements_by_css_selector("tbody>tr")[i].find_elements_by_tag_name("td")
                td_status = td_list[7].text #第八列：状态
                statusvalue = "审核通过" if (randvalue == 0) else "审核不通过"
                self.assertEqual(td_status, statusvalue, "状态与审核时勾选的不一致")

            except Exception as e:
                print("◆◆◆◆出现异常了◆◆◆◆ \r\n :" + str(e))
                continue
        time.sleep(10)

    def editexam(self):
        print("开始执行editNews()")
        driver = self.driver
        driver.get(self.base_url)
        login(driver)

        # 在平台页面，点击“新闻管理”
        driver.find_element_by_css_selector("div[class='item-main item-5']").click()
        time.sleep(1)

        # 初始化Excel对象，用来接收传过来的Json
        excelobj = Util.Excel("F:\新闻管理系统\Test Data\新闻管理.xlsx")
        sheetdata = excelobj.readSheet("添加新闻")

        news_dict = dict()
        for i in range(0, len(sheetdata)):
            news_dict[sheetdata[i]["新闻标题"].strip()] = sheetdata[i]["正文"].strip()

        # print(news_dict.__str__())

        for i in range(0, len(driver.find_elements_by_css_selector("tbody>tr"))):
            try:
                driver.find_elements_by_css_selector("tbody>tr>td:nth-child(9)>div>a[title='修改']")[i].click()
                print("点开了编辑按钮")
                time.sleep(2)

                self.assertEqual(driver.find_element_by_css_selector("#edit-modal-content>div.modal-header>h4").text,
                                 "编辑新闻",
                                 "编辑新闻窗口弹出失败，有BUG！")
                # “编辑新闻”弹出的窗体，文本框内的标题title_editnews，与外面的标题进行比对
                rand_flag = random.randint(0,1) #有概率的去编辑新闻标题
                if rand_flag == 0:
                    driver.find_element_by_id("newsTitle").clear()
                    driver.find_element_by_id("newsTitle").send_keys("哈哈哈哈哈哈哈")
                else:
                    pass
                title_editnews = driver.find_element_by_id("newsTitle").text #编辑新闻，文本框内的新闻标题

                td_list = driver.find_elements_by_css_selector("tbody>tr")[i].find_elements_by_tag_name("td")
                td_newstitle = td_list[2].text #第三列：新闻标题


                driver.find_element_by_css_selector("button[onclick='saveNewsInfo()']").click()
                time.sleep(1)

                # 弹出“操作完成”提示，点确定
                driver.find_element_by_xpath("/html/body/div[7]/div/div/div[3]/button").click()
                time.sleep(1)

                self.assertEqual(title_editnews, td_newstitle, "编辑新闻的新闻标题与列表中的新闻标题不一致")
            except Exception as e:
                print("◆◆◆◆出现异常了◆◆◆◆ \r\n :" + str(e))
                continue

    def deleteexam(self):
        print("开始执行deleteNews()")
        driver = self.driver
        driver.get(self.base_url)
        login(driver)

        # 在平台页面，点击“新闻管理”
        driver.find_element_by_css_selector("div[class='item-main item-5']").click()
        time.sleep(1)

        # 初始化Excel对象，用来接收传过来的Json
        excelobj = Util.Excel("F:\新闻管理系统\Test Data\新闻管理.xlsx")
        sheetdata = excelobj.readSheet("添加新闻")

        news_dict = dict()
        for i in range(0, len(sheetdata)):
            news_dict[sheetdata[i]["新闻标题"].strip()] = sheetdata[i]["正文"].strip()

        # print(news_dict.__str__())

        for i in range(0, len(driver.find_elements_by_css_selector("tbody>tr"))):
            try:
                td_list = driver.find_elements_by_css_selector("tbody>tr")[i].find_elements_by_tag_name("td")
                td_newstitle_before = td_list[2].text #第三列：新闻标题
                print("删除之前，新闻标题是：" + td_newstitle_before)


                driver.find_elements_by_css_selector("tbody>tr>td:nth-child(9)>div>a[title='删除']")[i].click()
                print("点击了删除按钮")
                time.sleep(2)

                self.assertEqual(driver.find_element_by_css_selector("body>div.bootbox.modal.fade.in>div>div>div.modal-header>h4").text,
                                 "系统提示",
                                 "系统提示窗体未弹出，有BUG！")

                # 弹出“系统提示”，是否删除新闻？rand_flag = 1，点击删除 ； rand_flag = 2 点击取消
                rand_flag = random.randint(1, 2) #有概率的去删除新闻标题
                driver.find_element_by_xpath("/html/body/div[7]/div/div/div[3]/button[" + str(rand_flag) + "]").click()
                time.sleep(1)
                if rand_flag == 1:
                    # 弹出“系统提示”提示“已删除”，点确定
                    self.assertEqual(driver.find_element_by_xpath("/html/body/div[7]/div/div/div[2]/div").text,
                                     "删除新闻成功",
                                     "新闻删除时出异常")
                    driver.find_element_by_xpath("/html/body/div[7]/div/div/div[3]/button").click()
                    time.sleep(1)

                    #然后判断news_dict[新闻标题]，这个key是否存在
                    self.assertEqual(news_dict.__contains__(td_newstitle_before), False, "'确认删除'失败，新闻标题还在那里啊！")
                else:
                    #然后判断news_dict[新闻标题]，这个key是否存在
                    self.assertEqual(news_dict.__contains__(td_newstitle_before), True, "'取消删除'失败，新闻已被删，标题都找不到了！")

            except Exception as e:
                print("◆◆◆◆出现异常了◆◆◆◆ \r\n :" + str(e))
                continue
        time.sleep(3)
        driver.quit(self)
        ExamManager.deleteExam(self)

    def tearDown(self):
        self.driver.quit()
        self.assertEqual([], self.verificationErrors)


if __name__ == "__main__":
    suite = unittest.TestSuite()
    # suite.addTest(ExamManager("login"))
    # suite.addTest(ExamManager("addExam"))
    suite.addTest(ExamManager("addpaper"))

    runner = unittest.TextTestRunner(failfast=False)
    runner.run(suite)