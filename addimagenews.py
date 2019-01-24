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
import os
import Util

url_news = "http://59.175.153.36:10300/news/a/news/base?menuId=31"
username = "admin"
password = "123456"

if __name__ == "__main__":
	driver = webdriver.Chrome()
	driver.maximize_window()

	driver.get(url_news)
	bootstrap = Util.BootStrap(driver)
	bootstrap.login(url_news, username, password)
	bootstrap.goto("添加新闻")

	img_url = "http://www.xinhuanet.com/politics/xxjxs/2018-11/16/129995539_15423297918061n.jpg"
	# 新闻标题
	driver.find_element_by_id("newsTitle").send_keys("这是新闻标题")

	# 正文内容
	driver.find_element_by_xpath("//*[contains(@id,'toolbar-')]/div[15]/i").click()
	time.sleep(2)
	driver.find_element_by_xpath("//li[text()='网络图片']/parent::*/following-sibling::*/descendant::input").send_keys(img_url)  # 输入图片地址
	time.sleep(3)  # 单步调试，百分百成功插入。但是实际执行时，间隔时间要长一点，不然图片在富文本里加载不出来
	driver.find_element_by_xpath("//li[text()='网络图片']/parent::*/following-sibling::*/descendant::button[text()='插入']").click()  # 插入
	time.sleep(3)
	driver.find_element_by_xpath("//*[contains(@id,'text-elem')]").send_keys("新闻正文")

	# 信息来源
	driver.find_element_by_id("sources").send_keys("人民网")
	time.sleep(1)

	# 点击“确定”
	driver.find_element_by_xpath("//button[@onclick='saveNewsInfo()']").click()


	# os.system('taskkill /im chrome.exe /F')  # 关闭整个浏览器
	os.system('taskkill /im chromedriver.exe /F')  # 关闭整个浏览器，并杀掉chromedriver.exe进程
