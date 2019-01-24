# -*- coding: utf-8 -*-
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import NoAlertPresentException
import unittest, time, re, os
import xlrd, xlwt, xlutils, openpyxl
import json
import requests
from enum import Enum, unique
import traceback
import xlwings
import win32clipboard
import win32con
import pyautogui
from win32api import GetSystemMetrics  # 获取屏幕宽度高度
from itertools import *
from operator import itemgetter

DELAY = 0.3

def get_continuous_number(_raw_list):
    '''
    # https://blog.csdn.net/weixin_42555131/article/details/82020064
    解析一个int数组，如果有连续的，则用connector字符连接
    如： list = [1,2,3,4,8,9,10,12,18]
    返回['1:4', '8:10', 12, 18]
    :param _raw_list: 原始int数组
    :param connector: 连接器，默认是冒号(供Excel里的Range调用)，也可以根据需求改成其他的
    :return: 简化了的数组
    '''
    fun = lambda x: x[1]-x[0]  # 比较后一位和前一位是否相差一个数值，比如相差1
    index_list = []
    for k, g in groupby(enumerate(_raw_list), fun):
        list_temp = [j for i, j in g]  # 连续的数字列表
        if len(list_temp) > 1:
            #scope = str(min(list_temp)) + connector + str(max(list_temp))  # 用:来连接，方便Excel调用，这种写法不好用，涉及类型转换烦死人
            scope = (min(list_temp), max(list_temp))  # 元祖大法好！！
        else:
            scope = list_temp[0]
        index_list.append(scope)

    _list = index_list.copy()  # 创建一个副本出来，以免干扰了传入的raw_list
    for i in range(len(_list)):
        if i == 0:
            _list[0] = _list[0]
        else:
            if isinstance(_list[i], tuple):
                delta = _list[i][1] - _list[i][0]
                if isinstance(_list[i-1], tuple):
                    _list[i] = (_list[i-1][1] + 1, _list[i-1][1] + 1 + delta)
                else:
                    _list[i] = (_list[i-1]+1, _list[i-1]+1 + delta)
            else:
                if isinstance(_list[i-1], tuple):
                    _list[i] = _list[i-1][1] + 1
                else:
                    _list[i] = _list[i-1]+1
    return index_list, _list


def move_item_to_end(_list, _item_value):
    '''
    将不存在重复项的list里的某个值，移动到最后，
    如[1,2,3,4,6,7] 将6移动到最后[1,2,3,4,7,6]
    如['品牌', 'CPU', '硬盘', '价格', '显卡'] 将价格移动到最后'品牌', 'CPU', '硬盘', '显卡', '价格']
    :param item_value: 待移动的值
    :return: 移动后的list
    '''
    # 将不存在重复项的list里的某个值，移动到最后
    for _value in _list:
        if _value == _item_value:
            _index = _list.index(_value)
            _list.append(_value)
            _list.remove(_list[_index])
    return _list



def getclipboardtext():
    """
    python读取剪切板内容
    :return:
    """
    win32clipboard.OpenClipboard()
    data = win32clipboard.GetClipboardData(win32con.CF_TEXT)
    win32clipboard.CloseClipboard()
    return data.decode('gbk')


class Web():
    def __init__(self, _driver):
        self.__driver = _driver

    def _get_max_width(self, _elements):
        """
        只要是WebElement集合，就返回这些WebElement的最大宽度
        :param _elements: WebElement的实例集合
        :return: 返回一个元素的最大宽度
        """
        elements_width = []
        for i in range(len(_elements)):
            elements_width.append(_elements[i].rect["width"])
        max_width = max(elements_width)
        return max_width

    def copy_rich_text(self, elements):
        """
        获取富文本，并复制到剪切板
        :param _driver:  webdriver实例
        :param xpath_div: xpath
        :return:
        """
        screen_height = GetSystemMetrics(1)  # 屏幕高度

        driver = self.__driver

        max_width = self._get_max_width(elements)

        start = elements[0]
        start_height = start.rect["height"]
        start_width = start.rect["height"]

        offsetx = 0  # 废弃
        offsety = 0  # 废弃  pyautogui的原点是(0,0)，而页面原点可能是(0,120)，解决办法：模拟F11 + 设置浏览器参数chrome_option "disable-infobars"

        srolltime = 15  # TODO
        try:
            pyautogui.FAILSAFE = False  # 禁用此特性，否则会抛异常FailSafeException
            pyautogui.moveTo(start.location["x"] + offsetx, start.location["y"] + offsety, duration=2)

            pyautogui.mouseDown(button="left", duration=1)

            pyautogui.moveTo(start.location["x"] + offsetx + max_width, screen_height, duration=2)
            time.sleep(srolltime)

            end = elements[len(elements) - 1]  # 最后一个元素
            end_height = end.rect["height"]  # 最后一个元素的高度
            end_width = end.rect["width"]  # 最后一个元素的宽度
            pyautogui.moveTo(start.location["x"] + max_width,
                             screen_height - divmod(end.location["y"] + end_height, screen_height)[1],
                             # 页面下滑到最下面之后，最后元素的高度
                             duration=2)  # 页面下滑到最下面之后，再移动到最后一个元素的右下角
            time.sleep(5)
            pyautogui.mouseUp(button="left", duration=0.1)
            time.sleep(1)
            pyautogui.hotkey("ctrl", "c")

        except Exception as e:
            traceback.print_exc()

        print(getclipboardtext())  # debug only


@unique
class TreeType(Enum):
    """
    树结构，两种类型的控件checkbox和radiobutton
    """
    CHECKBOX = 1
    RADIOBUTTON = 2


@unique
class TreeNodeLevel(Enum):
    """
    树节点类型
    """
    ALL = 1  # 所有
    LEVEL0 = 2  # 一级节点
    LEVEL1 = 3  # 二级节点


class DateTime(object):
    def __init__(self, _datetime):
        self.__datetime = _datetime

    def getspliter(self):
        """
        得到年份分割器。默认写成"/"， TODO : 其实还有"-"，等分割。针对“2020年1月10日”这种字符，则没法得到分割器
        :return:  默认分割器"/"
        """
        return "/"

    @property
    def getYear(self):
        datetime = self.__datetime
        return datetime.split(DateTime.getspliter(self))[0]

    @property
    def getMonth(self):
        datetime = self.__datetime
        return datetime.split(DateTime.getspliter(self))[1]

    @property
    def getDay(self):
        datetime = self.__datetime
        return datetime.split(DateTime.getspliter(self))[2]

    def selectdatetime(self, driver):
        year = self.getYear
        month = self.getMonth
        day = self.getDay
        # 处理日历控件：
        # 1.先输入年份
        driver.find_element_by_xpath("//div[@class='menuSel YMenu']/following-sibling::input[@class='yminput']").click()
        driver.implicitly_wait(5)
        driver.find_element_by_xpath(
            "//div[@class='menuSel YMenu']/following-sibling::input[@class='yminputfocus']").clear()
        driver.find_element_by_xpath(
            "//div[@class='menuSel YMenu']/following-sibling::input[@class='yminputfocus']").send_keys(year)
        time.sleep(1)

        # 2.输入月份
        driver.find_element_by_xpath("//div[@class='menuSel MMenu']/following-sibling::input[@class='yminput']").click()
        driver.implicitly_wait(5)
        driver.find_element_by_xpath(
            "//div[@class='menuSel MMenu']/following-sibling::input[@class='yminputfocus']").clear()
        driver.find_element_by_xpath(
            "//div[@class='menuSel MMenu']/following-sibling::input[@class='yminputfocus']").send_keys(month)
        time.sleep(1)

        # 3.选择日期
        driver.find_element_by_id("dpTimeStr").click()
        driver.implicitly_wait(5)
        date = year + "," + month + "," + day
        driver.find_element_by_xpath("//td[@onclick='day_Click(" + date + ");']").click()
        time.sleep(1)

        # 4.点击确定
        print(driver.find_element_by_xpath("//input[@value='确定']").is_displayed())  # debug only
        driver.find_element_by_xpath("//input[@value='确定']").click()
        time.sleep(1)


class ZTree(object):
    """
    针对Bootstrap中的ZTree
    """

    def __init__(self, _driver):
        self.__driver = _driver

    def _getindex(self, value, nodetype=TreeNodeLevel.ALL):
        """
        获取索引
        :param value: 查找的目标文本值
        :param nodetype: 节点类型：默认是TreeNodeLevel.ALL
        :return: 树节点中，某个文本所在树节点的索引值
        """
        _driver_ = self.__driver

        if nodetype == TreeNodeLevel.LEVEL0:
            # 找到带有 [+] [-] 图标的节点，其中class不能包含docu
            treenodes = _driver_.find_elements_by_xpath(
                "/li[contains(@class,'level0')]/span[contains(@class,'level0')]/following-sibling::a")
            for i in range(0, len(treenodes)):
                if value == treenodes[i].get_attribute("title"):
                    return i
                    break

        elif nodetype == TreeNodeLevel.LEVEL1:
            # 找到 不带[+] [-] 图标的节点
            treenodes = _driver_.find_elements_by_xpath(
                "/li[contains(@class,'level1')]/span[contains(@class,'level1')]/following-sibling::a")
            for i in range(0, len(treenodes)):
                if value == treenodes[i].get_attribute("title"):
                    return i
                    break
        else:
            # nodetype == TreeNodeLevel.ALL:
            # 找到ZTree的所有节点
            treenodes = _driver_.find_elements_by_xpath("//span[contains(@class,'switch')]/following-sibling::a")
            for i in range(0, len(treenodes)):
                if value == treenodes[i].get_attribute("title"):
                    return i
                    break

    def _getvalues(self, nodetype=TreeNodeLevel.ALL):
        """
        获取树形结构的所有文本值，输出成list
        :param nodetype: 节点类型
        :return: 树节点所有文本值的list
        """
        _driver_ = self.__driver
        treenodesvalue = []  # 初始化一个空list

        if nodetype == TreeNodeLevel.LEVEL0:
            # 找到带有 [+] [-] 图标的节点，其中class不能包含docu
            treenodes = _driver_.find_elements_by_xpath(
                "//span[contains(@class,'switch')][not(contains(@class,'docu'))]/following-sibling::a")
            for i in range(0, len(treenodes)):
                treenodesvalue.append(treenodes[i].text)
            return treenodesvalue

        elif nodetype == TreeNodeLevel.LEVEL1:
            # 找到 不带[+] [-] 图标的节点
            treenodes = _driver_.find_elements_by_xpath(
                "//span[contains(@class,'switch')][contains(@class,'docu')]/following-sibling::a")
            for i in range(0, len(treenodes)):
                treenodesvalue.append(treenodes[i].text)
            return treenodesvalue

        else:
            # nodetype == TreeNodeLevel.ALL:
            # 找到ZTree的所有节点
            treenodes = _driver_.find_elements_by_xpath("//span[contains(@class,'switch')]/following-sibling::a")
            for i in range(0, len(treenodes)):
                treenodesvalue.append(treenodes[i].text)
            return treenodesvalue

    def expandall(self):
        """
        展开所有树节点
        :return:
        """
        _driver_ = self.__driver

        # 找到带有 [+] [-] 图标的节点，其中class不能包含docu
        treenodes = _driver_.find_elements_by_xpath("//span[contains(@class,'switch')][not(contains(@class,'docu'))]")
        for i in range(0, len(treenodes)):
            if "close" in treenodes[i].get_attribute("class"):
                treenodes[i].click()
                time.sleep(0.5)
            else:
                continue

    def collideall(self):
        """
        收齐所有树节点
        :return:
        """
        _driver_ = self.__driver

        # 找到带有 [+] [-] 图标的节点，其中class不能包含docu
        treenodes = _driver_.find_elements_by_xpath("//span[contains(@class,'switch')][not(contains(@class,'docu'))]")
        for i in range(0, len(treenodes)):
            if "open" in treenodes[i].get_attribute("class"):
                treenodes[i].click()
                time.sleep(0.5)
            else:
                continue

    def expandbyindex(self, index):
        """
        展开指定index的树节点
        :param index: 索引
        :return:
        """
        _driver_ = self.__driver

        # 找到带有 [+] [-] 图标的节点，其中class不能包含docu
        treenodes = _driver_.find_elements_by_xpath("//span[contains(@class,'switch')][not(contains(@class,'docu'))]")
        if "close" in treenodes[index].get_attribute("class"):
            treenodes[index].click()
        else:
            pass

    def collidebyindex(self, index):
        """
        收起指定index的树节点
        :param index: 索引
        :return:
        """
        _driver_ = self.__driver

        # 找到带有 [+] [-] 图标的节点，其中class不能包含docu
        treenodes = _driver_.find_elements_by_xpath("//span[contains(@class,'switch')][not(contains(@class,'docu'))]")
        if "open" in treenodes[index].get_attribute("class"):
            treenodes[index].click()
        else:
            pass

    def checknodebyvalue(self, value):
        """
        根据文本值，勾选树节点
        :param value: 需要勾选的树节点的文本值
        :return:
        """
        _driver_ = self.__driver

        # 找到所有可勾选节点
        checkboxes = _driver_.find_elements_by_xpath("//span[contains(@class,'switch')]/following-sibling::span")
        try:
            index = self._getindex(value, TreeNodeLevel.ALL)
            if "true" in checkboxes[index].get_attribute('class'):  # 如果本来就是勾选的，则无需操作
                pass
            else:
                checkboxes[index].click()
        except Exception as e:
            traceback.print_exc()

    def unchecknodebyvalue(self, value):
        """
        根据文本值，反勾选树节点
        :param value: 需要反勾选的树节点的文本值
        :return:
        """
        _driver_ = self.__driver

        # 找到所有可勾选节点
        checkboxes = _driver_.find_elements_by_xpath("//span[contains(@class,'switch')]/following-sibling::span")
        try:
            index = self._getindex(value, TreeNodeLevel.ALL)
            if "false" in checkboxes[index].get_attribute('class'):  # 如果本来就是反勾选的，则无需操作
                pass
            else:
                checkboxes[index].click()
        except Exception as e:
            traceback.print_exc()

    def checkall(self, nodetype=TreeType.CHECKBOX):
        """
        选中所有树节点，特指checkbox
        :return:
        """
        _driver_ = self.__driver
        if nodetype == TreeType.CHECKBOX:
            checkboxes = _driver_.find_elements_by_xpath("//span[contains(@class,'check')]")
            for i in range(0, len(checkboxes)):
                if "true" in checkboxes[i].get_attribute('class'):  # 如果本来就是勾选的，则无需操作
                    pass
                else:
                    checkboxes[i].click()
                    time.sleep(DELAY)  # 考虑到相应速度，这里延迟DELAY = 0.3
        else:
            raise Exception("该类型的树结构，不支持全选")

    def uncheckall(self, nodetype=TreeType.CHECKBOX):
        """
        反选所有树节点，特指checkbox
        :return:
        """
        _driver_ = self.__driver
        if nodetype == TreeType.CHECKBOX:
            checkboxes = _driver_.find_elements_by_xpath("//span[contains(@class,'check')]")
            for i in range(0, len(checkboxes)):
                if "false" in checkboxes[i].get_attribute('class'):  # 如果本来就是反勾选的，则无需操作
                    pass
                else:
                    checkboxes[i].click()
                    time.sleep(DELAY)  # 考虑到相应速度，这里延迟DELAY = 0.3
        else:
            raise Exception("该类型的树结构，不支持全不选")

    def deselect(self, nodetype=TreeType.CHECKBOX):
        """
        反选所有树节点，特指checkbox
        :return:
        """
        _driver_ = self.__driver
        if nodetype == TreeType.CHECKBOX:
            checkboxes = _driver_.find_elements_by_xpath("//span[contains(@class,'check')]")
            for i in range(0, len(checkboxes)):
                checkboxes[i].click()
                time.sleep(DELAY)  # 考虑到相应速度，这里延迟DELAY = 0.3
        else:
            raise Exception("该类型的树结构，不支持全不选")


class BootStrap(object):
    """
    基于Bootstrap的网页，具有一些共通性，提炼出来
    """

    def __init__(self, _driver):
        self.__driver = _driver

    def login(self, url, username, password="123456"):
        """
        :param url: 地址，也就是保密综合管理系统中某个子系统的地址
        :param username: 用户名
        :param password: 密码
        :return:
        """
        driver = self.__driver
        driver.get(url)
        driver.maximize_window()
        try:
            driver.switch_to.alert.accept()
        except NoAlertPresentException as e:
            # print(e)  # debug only
            pass
        driver.find_element_by_id("username").clear()
        driver.find_element_by_id("password").clear()
        driver.find_element_by_id("username").send_keys(username)
        driver.find_element_by_id("password").send_keys(password)
        driver.find_element_by_css_selector("fieldset>div[class='clearfix']>button").click()
        driver.implicitly_wait(20)

    # region 废弃，后面由xpath选择器代替
    def _getcssselectorindex(self, _ctrlclass, _menuname):
        """
        :param _ctrlclass 控件的class名称 sidebar， navbar
        :param _menuname: 左侧树菜单结构，如涉密信息系统管理的左侧
        :return: 返回css selector的索引号，从1开始
        """
        # 上方导航结构 div[class^='navbar']>ul>li
        # 左侧树结构 div[class^='sidebar']>ul>li

        # 导航文本内容 div[class^='navbar']>ul>li>a>div[class='menu-text']
        # 树结构文本内容 div[class^='side']>ul>li>a>span[class='menu-text']
        # 综合一下通用css选择器表达式： div[class^='side']>ul>li>a>*[class='menu-text']
        _driver = self.__driver
        _driver.maximize_window()
        menu_list = _driver.find_elements_by_css_selector("div[class^='" + str(_ctrlclass) + "']>ul>li")
        for _index in range(0, len(menu_list)):
            print(menu_list[_index].text)
            if menu_list[_index].text == _menuname:
                return _index + 1  # 由于css选择器里面的起点是1，跟数组/list的起点0相差1，这里补1

    def gotobarmenu(self, ctrlclass, menuname):
        _driver = self.__driver
        _driver.maximize_window()
        _index = self._getcssselectorindex(ctrlclass.value, menuname)
        ctrlclassvalue = ctrlclass.value
        selector = "div[class^='" + str(ctrlclassvalue) + "']>ul>li:nth-child(" + str(_index) + ")"
        _driver.find_element_by_css_selector(selector).click()

    # endregioin

    def goto(self, text):
        text.replace("&nbsp;", " ").replace("　", " ").replace(" ", "")  # 先替换转义的空格，再替换全角空格，最后将空格去掉
        """
        无论是
        (1)左侧的树结构，还是
        (2)页面上方的页签
        (3)功能按钮
        (4)radio button  这个不行：//input/following-sibling::label[starts-with(@for,'exam')]    这个可以：//input[following-sibling::label[starts-with(@*,'')]]
        (5)下拉选项Select>Option标签，(可能不适用，需要单独处理) → 见selectbytext(text)方法
        (6)树结构的checkbox，如组织架构、试卷范围等  → 见ZTree这个类
        //span[starts-with(@id,'ztree')][contains(@id,'check')]/following-sibling::a/span[contains(@id,'span')]
        都可以用通用表达式xpath查找：
        :param text: 控件名（菜单、按钮），或者下拉框选项的值
        :return:
        """
        _driver = self.__driver
        # 通用xpath选择器，如有变动，只需要更改这里
        xpath_menu_button = "//*[starts-with(@class,'btn btn-sm btn-danger')]  | //div[contains(@class, 'bar')]/ul/li/a | //input[following-sibling::label[starts-with(@*,'')]]"
        for j in range(0, len(_driver.find_elements_by_xpath(xpath_menu_button))):
            # print(_driver.find_elements_by_xpath(xpath_menu_button)[j].text) # debug only
            element = _driver.find_elements_by_xpath(xpath_menu_button)[j]

            # 先替换转义的空格，再替换全角空格，最后将空格去掉
            if element.text.replace("&nbsp;", " ").replace("　", " ").replace(" ", "") == text:
                element.click()
                break
        time.sleep(1)

    def _getselectindex(self, value):
        """
        获取value所在下拉选项的索引位置
        :param value: 值
        :return: 索引
        """
        driver = self.__driver
        options = driver.find_elements_by_tag_name("option")

        for i in range(0, len(options)):
            if options[i].text == value:
                return i

    def selectbytext(self, optiontext):
        """
        根据选项的文本值，直接选中该选项。
        一般来说先获取Select对象，_select = Select(driver.find_element_by_id(id))
        然后根据不同方式选择，比如文本 _select.select_by_visible_text(text)
        此写法需要传入id和text。不同的Select不同的id，如果页面有十几个下拉框，那么多id非常冗余一旦id有变化代码维护很麻烦
        通过xpath的方式，在定位Select标签的时候，将text作为判断条件，即“包含了text的这个Select”，
        这样整个函数只需要一个参数，就是text本身，不需要任何id
        :param optiontext: 选项的文本值
        :return:
        """
        driver = self.__driver
        # _xpath = "//select[option[text(),'" + optiontext + "')]]"    # 不严谨，文本后面包含空格，或者&nbsp;无法匹配
        _xpath = "//select[option[contains(normalize-space(text()),'" + optiontext + "')]]"
        try:
            Select(driver.find_element_by_xpath(_xpath)).select_by_visible_text(optiontext)
        except Exception:  # 异常处理，打印出出异常的具体信息
            traceback.print_exc()

    def get_thead(self):
        """
        获取某个页签的listview的thead
        :return: 返回一个thead的list集合
        """
        _driver = self.__driver
        _titlelist = []
        _len_thead = len(_driver.find_elements_by_css_selector("table#sample-table-1>thead>tr>th"))
        for i in range(0, _len_thead - 1):
            if i == 0 or i == _len_thead - 1:  # 跳过首尾
                continue
            else:
                _titlelist.append(_driver.find_elements_by_css_selector("table#sample-table-1>thead>tr>th")[i].text)
        # print(_titlelist)
        return _titlelist

    def get_tbody(self):
        """
        获取某个页签的listview的tbody
        :return:  返回一个tbody的dict[list]符合json标准的结构
        """
        _driver = self.__driver
        _titlelist = []
        _width_tbody = len(_driver.find_elements_by_css_selector("table#sample-table-1>tbody>tr:nth-child(1)>td"))
        _height_tbody = len(_driver.find_elements_by_css_selector("table#sample-table-1>tbody>tr"))
        _dict = []
        for i in range(0, _height_tbody - 1):
            _tr = _driver.find_elements_by_css_selector("table#sample-table-1>tbody>tr")[i]  # 每一行tr看成一个元素
            _tbodylist = []
            for j in range(0, _width_tbody - 1):
                if j == 0 or j == _width_tbody - 1:  # 跳过首尾
                    continue
                else:
                    _tbodylist.append(_tr.find_elements_by_css_selector("td")[j].text)  # 以tr为基准，去查找td
            # print(_tbodylist)  # debug only
            _dict.append(_tbodylist)
        # print(_dict) # debug only
        return _dict


class Excel:
    """
    处理Excel
    """

    def __init__(self, excelname):
        self.__excelname = excelname

    '''
    @property
    def excelname(self):
        return self.__excelname

    @excelname.setter
    def excelname(self, value):
        self.__excelname = value

    @excelname.deleter
    def excelname(self):
        del self.__excelname
    '''

    # 将指定的单个Sheet里的单元格内容，每一行读取成一个字典，最后放在一个list里
    def readsheet(self, sheetname):
        """
        读取指定名称的sheet内容
        :param sheetname:  sheet名称
        :return: 返回当前sheet页内的数据，以list形式返回
        """
        _workbook = xlrd.open_workbook(self.__excelname)
        sheet = _workbook.sheet_by_name(sheetname)

        # 获取行数
        rows = sheet.nrows
        # print("行数是：")
        # print(rows) #debug only，result：PASS

        # 获取第一行数据
        _titlelist = sheet.row_values(0)

        # 定义一个list存放所有非第一行数据
        dictlist = {}
        list = []
        for row in range(1, rows):
            rowvalues = sheet.row_values(row)
            # print("每行的数据是：")
            # print(rowValues) #debug only，result：PASS

            dictdata = dict(zip(_titlelist, rowvalues))  # 将title列作为Key，将其余的列作为Value，循环打包成字典
            list.append(dictdata)

        # print("readSheet()返回结果：" + str(dictlist))
        dictlist[sheet.name] = list  # 2018/12/5 修改代码，返回完整json格式
        return dictlist

    # 循环读取Excel表格里的所有Sheet里的单元格内容,最终返回一个符合json格式的dict、list嵌套结构
    def readsheets(self):
        """
        循环读取所有sheet页里的内容
        :return: 返回一个符合json规范的dict[list]嵌套结构
        """
        _workbook = xlrd.open_workbook(self.__excelname)
        _sheets = _workbook.sheet_names()  # 通过Excel的Workbook对象获取到Sheet集合

        # print(str(len(self.sheets))) # debug only
        dictlists = {}  # 最外层是空字典
        for sheetno in range(0, len(_sheets)):
            # 获取行数
            currentsheetname = _sheets[sheetno]  # 当前Sheet名称
            currentsheet = _workbook.sheet_by_name(currentsheetname)  # 当前Sheet对象

            # print (self.sheets[sheetno]) #debug only
            rows = currentsheet.nrows  # 当前的Sheet里有nrows行

            # 获取第一行数据，作为标题数据，也就是dict里的key
            _titlelist = currentsheet.row_values(0)

            # 定义一个list存放所有非第一行数据
            # 非第一行循环添加为value
            dictlist = []
            for row in range(1, rows):  # 从第2行开始
                rowvalues = currentsheet.row_values(row)
                # print("每行的数据是：")
                # print(rowValues) #debug only，result：PASS

                dictdata = dict(zip(_titlelist, rowvalues))  # 将title列作为Key，将其余的列作为Value，循环打包成字典
                dictlist.append(dictdata)

            # print(self.dictlist)  #debug only
            # return self.dictlist

            # self.dictlists.append(self.dictlist)
            dictlists[_sheets[sheetno]] = dictlist

        # print("readSheets()返回结果：" + str(dictlists))  # debug only
        return dictlists

    def _write_data_to_cells(self, _data):
        """
        将数据循环写入到Excel当前Sheet的单元格里
        [
            {
                "Thai": "สวัสดีครับ/สวัสดีค่ะ",
                "English": "how do you do",
                "Korean": "안녕하세요",
                "Japanese": "こんにちは",
                "Spanish": "Buenos dias"
            }
        ]
        :param _data: 数据
        :return:
        """
        _keys = list(_data[0].keys())
        _length = len(_keys)

        # 先写入标题
        for k in range(0, _length):
            # Range函数，操作当前工作Sheet的单元格，参数是元祖，索引值从1开始
            xlwings.Range((1, k + 1)).value = _keys[k]

        for i in range(0, len(_data)):
            for j in range(_length):
                # Range从1开始，并且标题栏已经将Range("A1:D1")写入，则需要从A2开始，所以i+2
                xlwings.Range((i + 2, j + 1)).value = _data[i][_keys[j]]

    def writesheet(self, _sheetdata):
        """
        将dict{list[dict]}的json结构的数据_data写到名称为_sheet_name的sheet里
        {
                 "外语学习": [
                                {
                                    "Thai": "สวัสดีครับ/สวัสดีค่ะ",
                                    "English": "how do you do",
                                    "Korean": "안녕하세요",
                                    "Japanese": "こんにちは",
                                    "Spanish": "Buenos dias"
                                }
                            ]
        }
        :param _sheetdata: 数据
        :return: 生成指定目录下、指定名称、指定sheet名称的Excel文件，以xlsx形式保存
        """
        # https://blog.csdn.net/weixin_40693324/article/details/78855302
        # 将Excel读取的数据写到目标Excel里

        app = xlwings.App(visible=True, add_book=False)
        book = app.books.add()

        # 从读取的数据里将sheet名称解析出来
        _sheet_name = list(_sheetdata.keys())[0]

        book.sheets.add(_sheet_name)  # 添加一个Sheet
        book.sheets.__delitem__("Sheet1")  # 删除Sheet1

        _value = list(_sheetdata.values())[0]

        self._write_data_to_cells(_value)

        book.save(self.__excelname)
        book.close()
        app.kill()

    def writesheets(self, _sheetdata):
        """
        将dict{list[dict]}的json结构的数据_data写到名称为_sheet_name的sheet里
        {
            "这是sheet2": [
                            {
                                "A": 1,
                                "B": 2,
                                "C": 3
                            },
                            {
                                "A": 11,
                                "B": 22,
                                "C": 33
                            },
                            {
                                "A": 111,
                                "B": 222,
                                "C": 333
                            },
                            {
                                "A": 1111,
                                "B": 2222,
                                "C": 3333
                            }
                        ],
            "外语学习": [
                            {
                                "Thai": "สวัสดีครับ/สวัสดีค่ะ",
                                "English": "how do you do",
                                "Korean": "안녕하세요",
                                "Japanese": "こんにちは",
                                "Spanish": "Buenos dias"
                            }
                        ]
        }
        :param _sheetdata: 数据
        :return: 生成指定目录下、指定名称、指定sheet名称的Excel文件，以xlsx形式保存
        """
        app = xlwings.App(visible=True, add_book=False)
        book = app.books.add()

        length = len(_sheetdata.keys())
        for m in range(0, length):
            sheet = list(_sheetdata.keys())[m]
            book.sheets.add(sheet, before="Sheet1")  # 加上before="Sheet1"，才会按顺序添加Sheet
            _value = list(_sheetdata.values())[m]
            # print(_value)

            for n in range(len(list(_sheetdata.values())[m])):
                self._write_data_to_cells(_value)

        book.sheets.__delitem__("Sheet1")  # 最后删除Sheet1

        book.save(self.__excelname)
        book.close()
        app.kill()


class Token:
    """
    处理Token
    """

    def __init__(self, url, data):
        self.__url = url
        self.__data = data

    @property
    def gettoken(self):
        url = self.__url
        data = self.__data
        data = json.dumps(data)
        headers = {
            'content-type': "application/json",
            'cache-control': "no-cache",
            'postman-token': "0fb7a29d-a287-fd6d-1101-748fa2e426d2"
        }

        headers2 = {
            "Cache-control": "no-cache",
            "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10.11; rv:46.0) Gecko/20100101 Firefox/46.0",
            "Accept": "application/json, text/javascript, */*; q=0.01",
            "Accept-Language": "zh-CN,zh;q=0.8,en-US;q=0.5,en;q=0.3",
            "Accept-Encoding": "gzip, deflate",
            "Content-Type": "application/json;charset=UTF-8",
            "X-Requested-With": "XMLHttpRequest",
            "Connection": "keep-alive",
            "Content-Length": "888",
            "charset": "UTF-8"
        }

        response = requests.request("POST", url, headers=headers2, data=data)
        return response.json()['token']

    # 获取request请求的返回值的json格式
    @property
    def getresponsedata(self):
        url = self.__url
        data = self.__data
        data = json.dumps(data)
        headers = {
            'content-type': "application/json",
            'cache-control': "no-cache",
            'postman-token': "0fb7a29d-a287-fd6d-1101-748fa2e426d2"
        }

        headers2 = {
            "Cache-control": "no-cache",
            "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10.11; rv:46.0) Gecko/20100101 Firefox/46.0",
            "Accept": "application/json, text/javascript, */*; q=0.01",
            "Accept-Language": "zh-CN,zh;q=0.8,en-US;q=0.5,en;q=0.3",
            "Accept-Encoding": "gzip, deflate",
            "Content-Type": "application/json;charset=UTF-8",
            "X-Requested-With": "XMLHttpRequest",
            "Connection": "keep-alive",
            "Content-Length": "888",
            "charset": "UTF-8"
        }

        response = requests.request("POST", url, headers=headers2, data=data)
        return response.json()


class Paging(object):
    """
    Bootstrap分页
    也可以在页面直接调用javascript，代码如下
        js = "onclick=page(100,10,'');"
        driver.execute_script(js)
    """

    def __init__(self, driver):
        self.__driver = driver
        self.__maxpage = self.getmaxpage
        self.__step = 10

    @property
    def step(self):
        return self.__step

    @step.setter
    def step(self, value):
        self.__step = value

    @step.deleter
    def step(self):
        del self.__step

    def gettotal(self):
        pageinfo = self.getpageinfo
        total = re.findall("\d+", pageinfo)[0]  # 正则表达式查询数字
        return int(total)

    def getmaxpage(self):
        driver = self.__driver
        maxnumber = len(driver.find_elements_by_css_selector("div.pagination>ul>li"))
        maxpage = driver.find_element_by_css_selector(
            "div.pagination>ul>li:nth-child(" + str(maxnumber - 2) + ")>a").text
        return maxpage

    def getpageinfo(self):
        driver = self.__driver
        linklength = len(driver.find_elements_by_css_selector("a[href='javascript:']"))
        pageinfo = driver.find_elements_by_css_selector("a[href='javascript:']")[linklength - 1].text
        return pageinfo

    '''
    点击"下一页"
    '''

    def movetonext(self):
        driver = self.__driver
        linklength = len(driver.find_elements_by_css_selector("a[href='javascript:']"))
        driver.find_elements_by_css_selector("a[href='javascript:']")[linklength - 2].click()
        driver.implicitly_wait(3)

    '''
    点击"上一页"
    '''

    def movetoprevious(self):
        driver = self.__driver
        driver.find_elements_by_css_selector("a[href='javascript:']")[0].click()
        driver.implicitly_wait(3)

    '''
    点击"下一页"按钮，直到最后一页
    '''

    def moveforward(self):
        driver = self.__driver
        for i in range(0, self.__maxpage - 1):
            # print("点击'下一页'按钮第" + str(i) + "次") #debug only
            linklength = len(driver.find_elements_by_css_selector("a[href='javascript:']"))
            driver.find_elements_by_css_selector("a[href='javascript:']")[linklength - 2].click()
            time.sleep(1)

    '''
    点击"上一页"按钮，直到第一页
    '''

    def movebackward(self):
        driver = self.__driver
        for i in range(0, self.__maxpage - 1):
            # print("点击'下一页'按钮第" + str(i) + "次") #debug only
            driver.find_elements_by_css_selector("a[href='javascript:']")[0].click()
            time.sleep(1)

    '''
    点击最后一页的数字链接
    '''

    def movetolastpage(self):
        driver = self.__driver
        linklen = len(driver.find_elements_by_css_selector("div.pagination>ul>li>a"))
        driver.find_element_by_css_selector("div.pagination>ul>li:nth-child(" + str(linklen - 2) + ")>a").click()
        time.sleep(1)

    '''
    点击第一页的数字连接
    '''

    def movetofirstpage(self):
        driver = self.__driver
        driver.find_elements_by_css_selector("a[href='javascript:']")[1].click()
        time.sleep(1)


if __name__ == "__main__":
    '''
    # 测试readsheet()
    _directory = u"F:\新闻管理系统\Test Data"
    _excelname = "新闻管理.xlsx"
    _sheetname = "添加新闻"
    _excelfullname = os.path.join(_directory, _excelname)

    excel = Excel(_excelfullname)
    excel.readsheet(_sheetname)
    print("====================================================================")
    # excel.readsheets()

    # 初始化Excel对象，用来接收传过来的Json
    excelpath = "F:\新闻管理系统\Test Data\新闻test.xlsx"
    excelobj = Excel(excelpath)
    sheetdata_dict = excelobj.readsheets()

    # 获取Excel所有的sheet名称
    workbook = openpyxl.load_workbook(excelpath)
    sheetname_list = workbook.sheetnames

    for sheetname in sheetname_list:
        print(sheetdata_dict[sheetname])
    '''

    ''' 测试token
    url = "http://192.168.30.212:8080/preInfoLeak/a/Login"
    data = {
        "login_name": "sp1",
        "password": "111111",
        "data": {
            "client_mac": "2C:FD:A1:58:4C:A0",
            "client_OS": "Microsoft Windows10",
            "client_machine_name": "???????",
            "client_version": "1.1.1.3",
            "client_machineNO": "ethernet_32773"
        }
    }
    # response = Token.getresponsedata(url=url, data=data)
    # token = Token.gettoken(url=url, data=data)
    response = Token(url, data).getresponsedata
    token = Token(url, data).gettoken

    print(response)
    print(token)

    print("++++++++++++++++++++++++++")
    # Chrome浏览器无界面的使用方法，先引用Options
    from selenium.webdriver.chrome.options import Options
    options = Options()
    options.add_argument('-headless')   # 设置参数：无头参数
    # options.add_argument('--disable-gpu')   #设置参数：禁止GPU
    driver = webdriver.Chrome(chrome_options=options)
    driver.get("http://192.168.0.22:8801/exam/a/quest/etQuestion?menuId=6ce679b937f846e091be4da3021c4ee9")
    cookie_list = driver.get_cookies()
    print(cookie_list)
    '''
    # 打开文件选择框
    import win32ui

    fspec = "Excel文件 (*.xlsx)|*.xlsx|Excel文件 (*.xls)|*.xls|所有类型文件 (*.*)|*.*"
    flags = win32con.OFN_OVERWRITEPROMPT | win32con.OFN_FILEMUSTEXIST
    dialog = win32ui.CreateFileDialog(1, None, None, flags, fspec)  # 1表示打开文件对话框
    dialog.SetOFNInitialDir("C:\\")  # 初始位置
    dialog.SetOFNTitle("请选择一个合法的Excel文件")  # 对话框的标题
    dialog.DoModal()
    filepath = dialog.GetPathName()

    # 初始化Excel对象，用来接收传过来的Json
    # excelobj = Excel("F:\新闻管理系统\Test Data\新闻管理(多个sheet).xlsx")
    excelobj = Excel(filepath)
    data = excelobj.readsheets()
    print("sheetdata的值是 \r\n" + str(data))  # debug only

    # 另存为对话框
    dialog = win32ui.CreateFileDialog(0, None, None, flags, fspec)  # 0表示另存为对话框
    dialog.SetOFNInitialDir("C:\\")
    dialog.DoModal()
    _path = dialog.GetPathName()

    excel = Excel(_path)
    excel.writesheets(data)


