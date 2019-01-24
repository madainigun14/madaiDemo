# -*- coding: utf-8 -*-
from selenium import webdriver
import time
import traceback
import os
import Util

url_source = "http://www.xinhuanet.com/politics/xxjxs/2018-11/16/c_129995539.htm"  # 新闻来源
url_news = "http://59.175.153.36:10300/news/a/news/base?menuId=31"  # E路学习
url_news_1 = ""
username = "admin"
password = "123456"


if __name__ == "__main__":
	driver = webdriver.Chrome()
	driver.get(url_source)
	driver.maximize_window()

	# 新闻标题的xpath
	# 人民网、中国军网、新浪网、光明日报、中青在线、环球网、澎湃新闻、中央电视台 //h1[text()!='']
	# 新华网 //div[contains(@class,'title')][not(contains(@class,'hide'))]

	driver.implicitly_wait(30)
	news_title = driver.find_element_by_xpath("//h1[text()!='']| //div[contains(@class,'title')][not(contains(@class,'hide'))]").text
	print(news_title)

	# 新闻正文，主要类型有：
	# 段落 //div[count(node())>=3]/p
	# 文字链接 //p/descendant::a
	# 图片链接 //p/descendant::img
	# 视频 embed|video
	raw_text_list = driver.find_elements_by_xpath("//div[count(node())>=10]/p|//p/descendant::a|//p/descendant::img|embed|video")  # 找到段落、链接、图片、视频
	news_contents_list = []  # 空集合，用来装数据
	connector = "="
	for i in range(0, len(raw_text_list)):

		# 【处理】加粗<strong></strong>或<B></B>标签
		for index_strong in range(len(driver.find_elements_by_xpath("//strong"))):
			if raw_text_list[i] == driver.find_elements_by_xpath("//p[descendant::strong]")[index_strong]:  # 所有子节点中有strong的p
				p_text = raw_text_list[i].text
				# print("raw_Text_list[i].__dict__的值是： " + p_text)  # 标签p内的text，包括其子节点 debug only
				strong_text = driver.find_elements_by_xpath("//strong")[index_strong].text
				# print("strong_text的值是" + strong_text)  # 标签 strong内的text

				# 拆分p的text，分割器为strong的text
				s0 = str.split(p_text, strong_text)[0]
				s1 = "<strong>" + connector + strong_text
				s2 = str.split(p_text, strong_text)[1]

				news_contents_list.append(s0)
				news_contents_list.append(s1)
				news_contents_list.append(s2)
				continue

		for index_a in range(len(driver.find_elements_by_xpath("//p/descendant::a"))):
			if raw_text_list[i] == driver.find_elements_by_xpath("//p/descendant::a")[index_a]:  # 所有子节点中有a的p
				p_text = raw_text_list[i].text  # 此处和上面的循环里的p_text不一样，因为索引可能不一样
				a_text = driver.find_elements_by_xpath("//p/descendant::a")[index_a].text
				'''
				p0 = driver.find_elements_by_xpath("//p/descendant::a/parent::*/preceding-sibling::*[last()]")[1].text
				p1 = raw_text_list[i].text + connector + raw_text_list[i].get_attribute("href")
				p2 = driver.find_elements_by_xpath("//p/descendant::a/parent::*/following-sibling::*[last()]")[0].text
				'''
				p0 = str.split(p_text, a_text)[0]
				p1 = "<a>" + connector + a_text
				p2 = str.split(p_text, a_text)[0]

				news_contents_list.append(p0)
				news_contents_list.append(p1)  # 添加href值到集合里
				news_contents_list.append(p2)
				continue
		news_contents_list.append(raw_text_list[i].text + "\n")

		# 【处理】a
		# 条件0：所有p标签下包含的a标签，并且有text()值，语句为：//p/descendant::*[contains(name(),'a')][text()!='']
		# 条件1：所有包含了text()属性不为空的a标签的p标签，语句为：//p[descendant::*[contains(name(),'a')][text()!='']]
		if raw_text_list[i] == driver.find_element_by_xpath("//p[descendant::*[contains(name(),'a')][text()!='']]"):  # 如果集合的标签与条件1标签一致，则不要添加，直接跳过
			continue

		# 【处理】strong
		if raw_text_list[i] == driver.find_elements_by_xpath("//p[descendant::strong]"):
			continue

		# 【处理】图片img src=''
		if raw_text_list[i].get_attribute("src") != None:  # 如果找到的元素包含了src
			news_contents_list.append(raw_text_list[i].get_attribute("src"))  # 添加src值到集合里
			continue
		print(news_contents_list)
	print(news_contents_list)   # debug

	driver.execute_script('window.open(arguments[0])', url_news)
	handles = driver.window_handles
	driver.switch_to.window(handles[1])

	driver.get(url_news)
	bootstrap = Util.BootStrap(driver)
	bootstrap.login(url_news, username, password)
	bootstrap.goto("添加新闻")
	time.sleep(5)

	# 插入超链接按钮的xpath   //*[@id="toolbar-elem8484681478212517"]/div[10]/i   //*[contains(@id,'toolbar-')]/div[10]/i

	# 新闻标题
	driver.find_element_by_id("newsTitle").send_keys(news_title)
	driver.implicitly_wait(5)

	# 正文内容
	for j in range(0, len(raw_text_list)):
		if "http" in news_contents_list[j]:
			if "jpg" in news_contents_list[j] or "png" in news_contents_list[j]:  # 如果找到的元素是图片，jpg，png
				driver.find_element_by_xpath("//*[contains(@id,'toolbar-')]/div[15]/i").click()  # 点击【插入图片】按钮
				time.sleep(0.3)
				driver.find_element_by_xpath("//li[text()='网络图片']/parent::*/following-sibling::*/descendant::input").send_keys(news_contents_list[j])  # 输入图片地址
				driver.find_element_by_xpath("//li[text()='网络图片']/parent::*/following-sibling::*/descendant::button[text()='插入']").click()  # 插入
				time.sleep(2)
				driver.find_element_by_xpath("//*[contains(@id,'text-elem')]").send_keys("\n")  # 只要插入了图片，就换行
				continue
			else:
				driver.find_element_by_xpath("//*[contains(@id,'toolbar-')]/div[10]/i").click()  # 点击【插入超链接】按钮
				time.sleep(0.3)

				# 拆分一下字符串
				text = str.split(news_contents_list[j], connector)[0]
				href = str.split(news_contents_list[j], connector)[1]
				driver.find_elements_by_xpath("//li[text()='链接']/parent::*/following-sibling::*/descendant::input")[0].send_keys(text)  # 输入链接文字
				driver.find_elements_by_xpath("//li[text()='链接']/parent::*/following-sibling::*/descendant::input")[1].send_keys(href)  # 输入链接地址
				driver.find_element_by_xpath("//li[text()='链接']/parent::*/following-sibling::*/descendant::button[text()='插入']").click()  # 插入
				continue
		elif "<strong>" in news_contents_list[j]:
			driver.find_element_by_xpath("//*[contains(@id,'toolbar-')]/div[2]/i").click()  # 点击【加粗】按钮
			text = str.split(news_contents_list[j], connector)[1]
			driver.find_element_by_xpath("//*[contains(@id,'text-elem')]").send_keys(text)
			driver.find_element_by_xpath("//*[contains(@id,'toolbar-')]/div[2]/i").click()  # 再次点击【加粗】按钮，复原
			continue

		else:
			driver.find_element_by_xpath("//*[contains(@id,'text-elem')]").send_keys(news_contents_list[j])

	# 信息来源
	driver.find_element_by_id("sources").send_keys("人民网")

	# 点击“确定”
	driver.find_element_by_xpath("//button[@onclick='saveNewsInfo()']").click()
	time.sleep(1)

	#
	driver.find_element_by_xpath("//button[text()='确定']").click()


	# os.system('taskkill /im chrome.exe /F')  # 关闭整个浏览器
	os.system('taskkill /im chromedriver.exe /F')  # 关闭整个浏览器，并杀掉chromedriver.exe进程
