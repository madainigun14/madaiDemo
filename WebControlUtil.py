# -*- coding: utf-8 -*-
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import NoAlertPresentException
import unittest, time, re, os
import xlrd, xlwt, openpyxl
import json


class Calender:
    def __init__(self, driver):
        self.__driver = driver

    def startdate(self):
        # 开始时间
        driver = self.__driver
        driver.find_element_by_id("startDate").click()
        time.sleep(1)
        iframe_startdate = driver.find_element_by_css_selector("body>div:nth-child(13)>iframe")
        driver.switch_to.frame(iframe_startdate)  # 切到iframe
        driver.find_element_by_id("dpTodayInput").click()  # 点击今天
        driver.switch_to.default_content()  # 切回
        time.sleep(0.5)

    def enddate(self):
        # 结束时间
        driver.find_element_by_id("endDate").click()
        time.sleep(1)
        iframe_enddate = driver.find_element_by_css_selector("body>div:nth-child(13)>iframe")
        driver.switch_to.frame(iframe_enddate)  # 切到iframe
        driver.find_element_by_css_selector("td[onclick='day_Click(2018,10,26);']").click()
        driver.find_element_by_id("dpOkInput").click()
        driver.switch_to.default_content()  # 切回
        time.sleep(0.5)


if __name__ == "__main__":
    url = "http://192.168.0.22:8801/inspection/client/taskManage/taskMainList?type=1&parentId=5ef39b1754d747b3994d59ebc71374f0"
    driver = webdriver.Chrome()
    driver.maximize_window()
    driver.implicitly_wait(10)
    driver.get(url)
    driver.find_element_by_id("username").clear()
    driver.find_element_by_id("password").clear()
    driver.find_element_by_id("username").send_keys("madai3")
    driver.find_element_by_id("password").send_keys("123456")
    driver.find_element_by_css_selector("fieldset>div[class='clearfix']>button").click()
    time.sleep(1)

    paging = Paging(driver)

    pageinfo = paging.getpageinfo
    print("pageinfo的值是：")
    print(pageinfo)
    print("++++++++++++++++++++")

    total = paging.gettotal
    print("total的值是：")
    print(total)
    print(type(total))
    print("++++++++++++++++++++")

    print("step的值是：")
    print(paging.step)
    paging.step = 20
    print("经过setter修改后，step的值是：")
    print(paging.step)
    del paging.step
    print("经过deleter删除后，step的值是：")
    try:
        print(paging.step)
    except Exception as e:
        print(e)
    paging.step = 11
    print("重新对step赋值，step的值是：")
    print(paging.step)
    print("++++++++++++++++++++")

    maxpage = paging.getmaxpage
    print(maxpage)
    print("===================")

    i = 0
    while i < 10:
        paging.moveforward()
        paging.movebackward()
        i = i+1

    time.sleep(5)
    driver.close()
    driver.quit()


