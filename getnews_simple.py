# -*- coding: utf-8 -*-
from selenium import webdriver
import time
import traceback
import os
import Util

url_source = "http://www.xinhuanet.com/politics/xxjxs/2018-11/16/c_129995539.htm"  # 新闻来源
url_news = "http://59.175.153.36:10300/news/a/news/base?menuId=31"  # E路学习
username = "admin"
password = "123456"


def appendvalue(list, value):
	if value == "":
		pass
	else:
		list.append(value)


if __name__ == "__main__":
	driver = webdriver.Chrome()
	driver.get(url_source)
	driver.maximize_window()

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
	raw_text_list = driver.find_elements_by_xpath(xpath)  # 一句话xpath，找到段落、链接、图片、视频
	news_contents_list = []  # 空集合，用来装数据
	connector = "="

	# 104
	num_total = len(driver.find_elements_by_xpath(xpath))

	# p包含子节点，且必须有text值，一共30个
	num_p_child = len(driver.find_elements_by_xpath("//p[count(*)!=0][text()]"))

	######

	for x in range(len(driver.find_elements_by_xpath(xpath))):
		xlist = driver.find_elements_by_xpath(xpath)
		# print(driver.find_elements_by_xpath(xpath)[x].tag_name)
		for y in range(len(driver.find_elements_by_xpath("//p[count(*)!=0][text()]"))):
			# print(driver.find_elements_by_xpath("//p[count(*)!=0][text()]")[y].text)
			ylist = driver.find_elements_by_xpath("//p[count(*)!=0][text()]")
			if xlist[x].text in ylist[y].text:
				if xlist[x].text == ylist[y].text:
					continue
				else:
					z = ylist[y].text[len(xlist[x].text):]
					print(z)
					continue
		print("===================")


	######

	for i in range(0, len(raw_text_list)):
		p_text = raw_text_list[i].text

		# p标签有以下几种情况：
		# (1).p只有文字
		# (2).p只有其子节点文字
		# (3).p有文字，其子节点也有文字
		# (4).p下面a有@href
		# (5).p下面img、embed、video有@src
		# 【处理】p元素下既包含text()，其子元素页包含text()的情况
		for index_p in range(len(driver.find_elements_by_xpath("//p[count(*)!=0][text()]"))):
			if raw_text_list[i] == driver.find_elements_by_xpath("//p[count(*)!=0][text()]")[index_p]:  # 如果找到了这种p
				p_self_text = driver.find_elements_by_xpath("//p[count(*)!=0][text()]")[index_p].text
				# print(p_self_text[len(p_text):])

		# 【处理】超链接，将标签a的text()和href属性都分别取出来，供富文本使用
		for index_a in range(len(driver.find_elements_by_xpath("//p/descendant::a[@href]"))):
			if raw_text_list[i] == driver.find_elements_by_xpath("//p[descendant::a[@href]]")[index_a]:  # 所有子节点中有a href的p
				a_text = driver.find_elements_by_xpath("//p/descendant::a")[index_a].text

				a0 = str.split(p_text, a_text)[0]
				a1 = "<a>" + connector + a_text
				a2 = str.split(p_text, a_text)[0]

				news_contents_list.append(a0)
				news_contents_list.append(a1)
				news_contents_list.append(a2)
				continue

		# 【处理】img,embed, video标签，src属性都分别取出来，供富文本使用
		for index_a in range(len(driver.find_elements_by_xpath("//p/descendant::*[@src]"))):
			if raw_text_list[i] == driver.find_elements_by_xpath("//p[descendant::*[@src]]")[index_a]:  # 所有子节点中有@src的p
				tag_name = driver.find_elements_by_xpath("//p/descendant::a")[index_a].tag_name
				tag_src = driver.find_elements_by_xpath("//p/descendant::a")[index_a].get_attribute("src")

				tag_info = tag_name + connector + tag_src

				news_contents_list.append(tag_info)
				continue

		# 【处理】强调<strong></strong>标签
		for index_strong in range(len(driver.find_elements_by_xpath("//strong"))):
			if raw_text_list[i] == driver.find_elements_by_xpath("//p[descendant::strong]")[index_strong]:  # 所有子节点中有strong的p
				strong_text = driver.find_elements_by_xpath("//strong")[index_strong].text

				# 拆分p的text，分割器为strong的text
				s0 = str.split(p_text, strong_text)[0]
				s1 = "<strong>" + connector + strong_text
				s2 = str.split(p_text, strong_text)[1]

				news_contents_list.append(s0)
				news_contents_list.append(s1)
				news_contents_list.append(s2)
				continue

		# 【处理】字体，也就是<font>标签，包括颜色，斜体，下划线
		font_elements = driver.find_elements_by_xpath("//*/font[text()]")
		for index_font in range(len(font_elements)):
			if raw_text_list[i] == driver.find_elements_by_xpath("//*/font[text()]")[index_font]:  # 所有子节点中有font的p
				# p_text = raw_text_list[i].text
				font_text = font_elements[index_font].text
				font_face = font_elements[index_font].get_attribute("face")
				font_size = font_elements[index_font].get_attribute("size")
				font_color = font_elements[index_font].get_attribute("color")

				f0 = str.split(p_text, font_text)[0]
				f1 = "<font>" + connector + font_text
				f2 = str.split(p_text, font_text)[1]

				news_contents_list.append(f0)
				news_contents_list.append(f1)
				news_contents_list.append(f2)

		news_contents_list.append(raw_text_list[i].text + "\n")
		# print(news_contents_list)  # debug only
	# print(news_contents_list)   # debug only

	# ---------------开始添加新闻---------------
	driver.execute_script('window.open(arguments[0])', url_news)  # 先执行javascript，在当前窗体，另起一个浏览器页签
	handles = driver.window_handles  # 获得当前浏览器页签的句柄
	driver.switch_to.window(handles[1])  # 切换到第2个句柄的窗体，也就是新开的页面

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
