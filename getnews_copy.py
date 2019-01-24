# -*- coding: utf-8 -*-
from selenium import webdriver
import time
import traceback
import os
import Util
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import win32clipboard
import win32con
import pyautogui
from win32api import GetSystemMetrics  # 获取屏幕宽度高度


url_source = "http://www.xinhuanet.com/politics/xxjxs/2018-11/16/c_129995539.htm"  # 新闻来源 (新闻网)
url_source_2 = "http://www.81.cn/xue-xi/2018-11/21/content_9351196.htm"  # 中国军网
url_source_3 = "https://www.thepaper.cn/newsDetail_forward_2648424"  # 澎湃新闻
url_news = "http://59.175.153.36:10300/news/a/news/base?menuId=31"  # E路学习
username = "admin"
password = "123456"


if __name__ == "__main__":
	option = webdriver.ChromeOptions()
	option.add_argument('disable-infobars')  # 关闭这个提示，以免影响网页坐标在屏幕的位置“Chrome正在受自动化控制”
	driver = webdriver.Chrome(chrome_options=option)
	driver.get(url_source)
	# driver.maximize_window()
	pyautogui.hotkey("F11")  # F11

	# 新闻标题的xpath
	# 人民网、中国军网、新浪网、光明日报、中青在线、环球网、澎湃新闻、中央电视台 //h1[text()!='']
	# 新华网 //div[contains(@class,'title')][not(contains(@class,'hide'))]

	driver.implicitly_wait(30)
	news_title = driver.find_element_by_xpath("//h1[text()!='']| //div[contains(@class,'title')][not(contains(@class,'hide'))]").text
	# 新闻正文，主要类型有：
	# 段落 //div[count(node())>=3]/p
	# 文字链接 //p/descendant::a
	# 图片链接 //p/descendant::img
	# 视频 embed|video
	xpath = "//p/descendant::img|embed|video[@src]|//p/descendant::a[text()]|//p/strong[text()]|//p/strong/font[text()]|//p/font[text()]|//div[count(node())>=10]/descendant::p[count(*)=0]|//p[count(*)!=0][text()]"
	xpath_div = "//div[@class='main-aticle']"
	elements = driver.find_elements_by_xpath(xpath_div)

	web = Util.Web(driver)
	web.copy_rich_text(elements)


	# ---------------开始添加新闻---------------
	driver.execute_script('window.open(arguments[0])', url_news)  # 先执行javascript，在当前窗体，另起一个浏览器页签
	handles = driver.window_handles  # 获得当前浏览器页签的句柄
	driver.switch_to.window(handles[1])  # 切换到第2个句柄的窗体，也就是新开的页面

	driver.get(url_news)
	bootstrap = Util.BootStrap(driver)
	bootstrap.login(url_news, username, password)
	bootstrap.goto("添加新闻")
	time.sleep(5)

	# 新闻标题
	driver.find_element_by_id("newsTitle").send_keys(news_title)
	driver.implicitly_wait(5)

	# 正文内容
	driver.find_element_by_xpath("//*[contains(@id,'text-elem')]").send_keys(Keys.CONTROL, "v")  # Ctrl + V
	time.sleep(10)

	# 信息来源
	driver.find_element_by_id("sources").send_keys("人民网")

	# 点击“确定”
	driver.find_element_by_xpath("//button[@onclick='saveNewsInfo()']").click()
	time.sleep(1)

	#
	driver.find_element_by_xpath("//button[text()='确定']").click()

	# os.system('taskkill /im chrome.exe /F')  # 关闭整个浏览器
	os.system('taskkill /im chromedriver.exe /F')  # 关闭整个浏览器，并杀掉chromedriver.exe进程
