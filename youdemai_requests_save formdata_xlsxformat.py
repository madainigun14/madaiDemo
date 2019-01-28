# -*- coding: utf-8 -*- python 3.6.5
import time, re, os
import xlrd, xlwt, xlutils, openpyxl
import json
import requests
from requests import adapters  # requests.adapters.DEFAULT_RETRIES = 5
import traceback
import xlwings
from urllib.parse import urlencode
from contextlib import closing
from lxml import etree
import urllib, urllib3
import gc, objgraph  # 监控内存
import threading
from time import sleep, ctime
from concurrent import futures  # 并发

'''
python 3.6.5
'''

MAX_WORKERS = os.cpu_count()  # 多进程依赖于“核”，这里获取当前计算机的Core数量
DELAY = 0.3  # 延迟
MIN_DELAY = 0.001
AMOUNT = 500  # 每次获得某个产品的多少条询价
requests.adapters.DEFAULT_RETRIES = 5  # 连接有问题时，重试5次
ITEM_PER_PAGE = 16  # 有得卖 > 笔记本电脑分页，每页固定展示16台机器
BASE_PRODUCT_URL = "http://www.youdemai.net"  # 有得卖官网
SORT_URL = "http://www.youdemai.net/products/product/sort"  # 查询某个品牌有多少待售电脑
MERLIST_URL = "http://www.youdemai.net/products/product/merlist" # 查询某个品牌的分页里每一页有多少电脑信息
BASE_QUERY_URL = 'http://www.youdemai.net/order/order/inquiry'  # 对某个待售电脑，以特定的参数组合查询价格

BRAND_DICT_EN={
    "thinkpad": "p0301",
    "apple": "p0308",
    "dell": "p0302",
    "hp": "p0303",
    "lenovo": "p0300",
    "asus": "p0304"
}

BRAND_DICT={
    "联想": "p0300",
    "ThinkPad": "p0301",
    "索尼": "p0305",
    "Apple": "p0308",
    "戴尔": "p0302",
    "惠普": "p0303",
    "asus": "p0304",
    "宏碁": "p0307",
    "小米": "p0651",
    "三星": "p0309",
    "微星": "p0319",
    "神舟": "p0311",
    "东芝": "p0312",
    "清华同方": "p0313",
    "海尔": "p0316",
    "富士通": "p0318",
    "技嘉": "p0320",
    "七喜": "p0321",
    "雷神": "p0335",
    "炫龙": "p0334",
    "机械革命": "p0333",
    "机械师": "p0461",
    "雷蛇": "p0332",
    "海鲅": "p0331",
    "Wbin": "p0330",
    "镭波": "p0327",
    "长城": "p0323",
    "明基": "p0322",
    "Alienware": "p0317",
    "方正": "p0315",
    "LG": "p0314",
    "微软": "p0310",
    "松下": "p0622",
    "其他": "p1004",
    "TaiAiCH": "p0642",
    "彗星人": "p0641",
    "天宝": "p0640",
    "维派": "p0639",
    "锋锐": "p0638",
    "宝扬": "p0637",
    "戴睿": "p0636",
    "dostyle": "p0635",
    "麦本本": "p0634",
    "AORUS": "p0633",
    "华为": "p0632",
    "火影": "p0631",
    "魔法师": "p0630",
    "幻影战士": "p0629",
    "NEC": "p0628",
    "万利达": "p0627",
    "优派": "p0626",
    "得峰": "p0625",
    "Gateway": "p0624",
    "主齿轮": "p0623",
    "七彩虹": "p0621",
    "AWO": "p0643",
    "未来人类": "p0650",
    "VOYO": "p0649",
    "ENZ": "p0648",
    "标逸": "p0647",
    "乐凡": "p0646",
    "紫麦": "p0644",
    "格莱富": "p0645"}

def get_brand_dict(_url="http://www.youdemai.net/products/product/brands?c=5&type=L&source="):
    brand_dict = {}
    _request_headers = {
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8",
        "Accept-Encoding": "gzip, deflate",
        "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
        "Connection": "close",
        "Host": "www.youdemai.net",
        "Upgrade-Insecure-Requests": "1",
        "User-Agent": "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/69.0.3497.100 Safari/537.36"
    }
    with requests.request("GET", _url) as response:
        # '<a href="/products/product/brands?c=5&amp;type=L&amp;pcode=p0300&amp;source=">联想</a>'
        html = etree.HTML(response.text)
        brand_name_list = html.xpath("//a[contains(@href, 'c=5&type=L&pcode=')]/text()")
        brand_list = html.xpath("//a[contains(@href, 'c=5&type=L&pcode=')]/@href")
        brand_code_list = []
        for i in range(len(brand_list)):
            brand_code_list.append(brand_list[i].split("&source")[0].split("pcode=")[1])

        print(brand_name_list)
        print(brand_code_list)
        for i in range(len(brand_name_list)):
            brand_dict[brand_name_list[i]] = brand_code_list[i]
    print(brand_dict)
    return brand_dict


def get_page_num(brand):
    num = get_url_list_count(brand)  # 通过post请求，解析返回的json得到thinkpad总共有多少台电脑出售
    page = divmod(num, ITEM_PER_PAGE)  # 返回的是一个(商，余数)的元祖

    # 返回有多少个分页
    if page[1] == 0:
        return page[0], 0  # 返回一个元祖 (多少个整页 , 最后一页有0个)
    else:
        return page[0] + 1, page[1]  # 返回一个元祖 (多少个整页 , 最后一页有几个)


def get_url_list_count(_brand):
    _request_headers = {
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8",
        "Accept-Encoding": "gzip, deflate",
        "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
        "Cache-Control": "max-age=0",
        "Connection": "close",
        "Content-Length": "438",
        "Content-Type": "application/x-www-form-urlencoded",
        "Host": "www.youdemai.net",
        "Origin": "http://www.youdemai.net",
        "Upgrade-Insecure-Requests": "1",
        "User-Agent": "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/69.0.3497.100 Safari/537.36"
    }

    _form_data = {
        "id": "",
        "keywords": "",
        "brandCode": BRAND_DICT[_brand],
        "brandName": "",
        "merType": "L",
        "merTypeName": "笔记本电脑"
    }
    _form_data = urlencode(_form_data)
    with requests.request("POST", SORT_URL, headers=_request_headers, data=_form_data) as response:

        return int(json.loads(response.text)[0]["LIST"][0]["SIZE"])
    _request_headers = {}  # 释放资源
    _form_data = {}  # 释放资源


def get_producturls_by_url(_brand):
    # 根据产品品牌的一个url，获取到待售所有电脑的具体地址，返回集合
    page = get_page_num(_brand)

    '''
    if isinstance(page, int):
        max_page_num = page
        last_page_count = 0
    if isinstance(page, tuple):
        max_page_num = page[0]
        last_page_count = page[1]
    '''
    max_page_num, last_page_count = page[0], page[1]  # 直接获取元祖

    product_url_list = []  # 返回所有电脑的具体地址集合
    spid_list = []  # 返回产品的一个form data请求里的特征值spid集合
    _request_headers = {
        "Accept": "*/*",
        "Accept-Encoding": "gzip, deflate",
        "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
        "Connection": "close",
        "Content-Length": "115",
        "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
        "Host": "www.youdemai.net",
        "Origin": "http://www.youdemai.net",
        "User-Agent": "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/69.0.3497.100 Safari/537.36",
        "X-Requested-With": "XMLHttpRequest"
    }
    for i in range(max_page_num):
        _form_data = {
            "id": "",
            "keywords": "",
            "brandCode": BRAND_DICT[_brand],
            "brandName": "",
            "merType": "L",
            "merTypeName": "笔记本电脑",
            "page": i+1
        }
        with requests.request("POST", MERLIST_URL, headers=_request_headers, data=_form_data) as response:
            json_text = json.loads(response.text)

            if last_page_count != 0 and i == max_page_num - 1:
                for k in range(page[1]):
                    spid = json_text["result"][k]["MERID"]
                    product_url = "http://www.youdemai.net/products/product/detail?id=" + spid+ "&source="
                    spid_list.append(spid)
                    product_url_list.append(product_url)
            else:
                for j in range(ITEM_PER_PAGE):
                    spid = json_text["result"][j]["MERID"]
                    product_url = "http://www.youdemai.net/products/product/detail?id=" + spid+ "&source="
                    spid_list.append(spid)
                    product_url_list.append(product_url)

    _request_headers = {}  # 释放资源
    _form_data = {}  # 释放资源
    return product_url_list, spid_list


def get_productname(_html):
    _product_name = _html.xpath("//div[@class='leftarea left']/div/div[text()][preceding-sibling::img]")[0].text
    product_name = _product_name.replace("/","_")
    return product_name

def get_productnames_from_producturls(_producturl_list):
    '''
    根据产品的链接地址，发送GET请求后，返回的网页，进行解析得到产品名称的集合
    :param _producturl_list: 产品链接地址集合
    :return: 产品名称集合
    '''
    product_name_list = []
    i = 1
    for _product_url in _producturl_list:
        with requests.request("GET", _product_url) as response:
            html = etree.HTML(response.text)
            productname = get_productname(html)
            product_name_list.append(productname)
            print("第" + str(i) + "个产品名：" + productname)
            i += 1
    print("产品名称集合获取完毕")
    return product_name_list


def get_incompleted_product_name(_brand, _product_name_list):
    PRODUCT_DIR_UTF8 = u"F:\\电脑报价\\youdemai\\" + _brand + "\\" + str(AMOUNT) + "\\utf-8\\"
    # 已完成
    product_completed_list = []
    product_names = os.listdir(PRODUCT_DIR_UTF8)
    for product_name in product_names:
        product_name = product_name.split(".txt")[0]
        product_completed_list.append(product_name)

    # 未完成
    product_incompleted_list =[]
    for _product_name in _product_name_list:
        if _product_name not in product_completed_list:
            product_incompleted_list.append(_product_name)
            print(_product_name)
    print(str(product_incompleted_list))
    return product_incompleted_list


out_list = []  # 一个数组用来装一堆二层数组get_arg_list递归时调用
out_code_list = []  # 一个数组用来装一堆二层数组get_arg_list递归时调用


def get_arg_list(html, param_list, x=0):
    '''

    :param html: 一个lxml.etree.HTML的对象实例，功能与Webdriver的driver一样html.xpath()等价于driver.find_elements_by_xpath()
    :param param_list: 一台电脑的参数的 大分类
    :param x:  递归函数的一个占位符，初始值为0
    :return:
    '''
    # xpath = "//li[div[div[text()='" + param_list[i] + "']]]/div/ul/li"
    # args = html.xpath("//li[div[div[text()]]]/div/ul/li")
    # i = 0  # 各个参数类型的子元素的序号
    i = x
    while True:
        xpath = "//li[div[div[text()='" + param_list[i] + "']]]/div/ul/li"  # param_list[i] =  CPU,xpath里需要加text()，不然会算多
        for j in range(1):
            # print("元素分别是")  # debug only
            inner_list = []  # 一个用来存text
            inner_code_list = []  # 一个用来存id
            # 用"//dl[contains(@data-value,'处理器')]/dd/ul/li/div" 查出来的N个元素
            for k in range(len(html.xpath(xpath))):
                # print(html.xpath(xpath)[k].text)  # debug only
                inner_list.append(html.xpath(xpath)[k].text)  # 添加要点击的元素，debug时添加.text属性，正式运行时添加元素本身
                inner_code_list.append(html.xpath(xpath)[k].attrib["id"])
                # inner_list.append(html.xpath(xpath)[k])  # 添加要点击的元素，debug时添加.text属性，正式运行时添加元素本身
            out_list.append(inner_list)  # 将内层数组加到最外层数组，形成二维数组
            out_code_list.append(inner_code_list)
        try:  # 必须放在try，finally结构里，不然递归的时候return不出去
            get_arg_list(html, param_list, x + 1)
        finally:
            return out_list, out_code_list


def get_formdata_and_price_by_url(_brand, _product_url):
    param_list = []  # formdata 参数分类的列表
    arg_text_combo_dict = {}  # formdata 参数的“文本”组合情况列表
    arg_code_combo_list = []  # formdata 参数的“编码”组合情况列表，并经过反转，#符号拼接
    _session = requests.session()  # 创建一个session


    with requests.request("GET", _product_url) as response:
        html = etree.HTML(response.text)
        elements_params = html.xpath("//div[@class='radio_title']")
        PRODUCTNAME = get_productname(html)
    response.close()
    # 循环将param写入param_list
    for i in range(len(elements_params)):
        param = elements_params[i].text
        #print(param)
        param_list.append(param)

    # 循环获取所有param下的arg，构造成二维数组
    arg_text, arg_code = get_arg_list(html, param_list, 0)
    cartesian_text = Cartesian(arg_text)
    cartesian_code = Cartesian(arg_code)

    arg_text_combo = cartesian_text.assemble()
    arg_code_combo = cartesian_code.assemble()

    param_list.append("价格")
    param_list.insert(0,"产品名称")
    param_list = "@".join(param_list)

    for j in range(len(arg_text_combo)):

        arg_code_combo[j].reverse()
        real_arg_code = "#".join(arg_code_combo[j])
        arg_code_combo_list.append(real_arg_code)

        # PRICE = get_price_by_session(_session, _product_url, real_arg_code)
        PRICE = get_price_by_post(_product_url, real_arg_code)
        time.sleep(DELAY)
        arg_text_combo[j].append(PRICE)
        arg_text_combo[j].insert(0, PRODUCTNAME)
        arg_text_combo[j] = "@".join(arg_text_combo[j])
        print("第" +str(j+1) + "次arg_text_combo获取完毕" + PRODUCTNAME)
        #print(arg_text_combo[j])
        if j % 500 == 0:  # 如果是500的倍数次
            time.sleep(5)  # 则休息5秒
        if j == AMOUNT:
            break  # 循环中断控制放在这里，先执行后判断是否中断。不然最后一条数据无法处理
    print(str(arg_text_combo))
    del arg_text_combo[AMOUNT:]  # 删除多余的，只留AMOUNT个数据
    arg_text_combo.insert(0, param_list)
    print(str(arg_text_combo))

    #处理jsresult_code数据，反转，“#”拼接
    #for i in range(len(arg_code_combo)):
    #    #if i == 50000: break
    #    arg_code_combo[i].reverse()
    #    temp = "#".join(arg_code_combo[i])
    #    arg_code_combo_list.append(temp)

    #jsonhelper = JsonHelper()
    #jsonhelper.write_json_to_txt(u"F:\\电脑报价\\youdemai\\thinkpad_param_price\\radioids\\" + PRODUCTNAME + "_请求参数radioids.txt", arg_code_combo_list)  # 字典转成json后再传参

    PRODUCT_DIR_UTF8 = u"F:\\电脑报价\\youdemai\\" + _brand + "\\" + str(AMOUNT) + "\\utf-8\\"
    PRODUCT_DIR_ANSI = u"F:\\电脑报价\\youdemai\\" + _brand + "\\" + str(AMOUNT) + "\\ansi\\"

    if not os.path.exists(PRODUCT_DIR_UTF8):  # 判断路径是否存在，如果不存在，递归创建文件夹
        os.makedirs(PRODUCT_DIR_UTF8)
    if not os.path.exists(PRODUCT_DIR_ANSI):  # 判断路径是否存在，如果不存在，递归创建文件夹
        os.makedirs(PRODUCT_DIR_ANSI)

    file2 = open(PRODUCT_DIR_UTF8 + PRODUCTNAME + ".txt", "w+", encoding="utf-8")
    for each in arg_text_combo: file2.write(each+"\r\n")
    #file2.write(str(arg_text_combo))
    file2.close()
    file3 = open(PRODUCT_DIR_ANSI + PRODUCTNAME + ".txt", "w+")
    for each in arg_text_combo: file3.write(each+"\r\n")
    #file2.write(str(arg_text_combo))
    file3.close()


    #excelpath = u"F:\\电脑报价\\youdemai\\thinkpad_param_price\\youdemai_post_raw_data\\excel_2\\" + PRODUCTNAME + ".xlsx"
    #excel = Excel(excelpath)

    # 先写入标题
    #for k in range(len(arg_text_combo)):
    #    # Range函数，操作当前工作Sheet的单元格，参数是元祖，索引值从1开始
    #   xlwings.Range((1, k + 1)).value = arg_text_combo[k]

    del arg_text[:]  # 【内存溢出】这个必须删除，不然下次循环到这个方法体，会累加
    del arg_code[:]  # 【内存溢出】这个必须删除，不然下次循环到这个方法体，会累加
    del cartesian_text
    del cartesian_code
    del arg_text_combo
    del arg_code_combo
    #del arg_code_combo_list[:]  # 删除对象防止内存溢出
    del param_list
    del _brand
    #del param_list[:]  # 删除对象防止内存溢出
    #category_param_dict.clear()
    #arg_text_combo_dict.clear()  # 删除对象防止内存溢出
    #del computer_param_list[:]  # 删除对象防止内存溢出


def get_price_by_post(_product_url, _radio_ids):
    _spid = _product_url.split("id=")[1].split("&")[0]  # _product_url与_spid本身存在一定关系
    _request_headers = {
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8",
        "Accept-Encoding": "gzip, deflate",
        "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
        "Cache-Control": "max-age=0",
        "Connection": "close",
        "Content-Length": "438",
        "Content-Type": "application/x-www-form-urlencoded",
        "Host": "www.youdemai.net",
        "Origin": "http://www.youdemai.net",
        "Referer": _product_url,
        "Upgrade-Insecure-Requests": "1",
        "User-Agent": "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/69.0.3497.100 Safari/537.36"
    }

    _form_data = {
        "spId": _spid,
        "source": "",
        "typeId": "",
        "radioIds": _radio_ids,
        "multiIds": "",
        "areaId": "",
        "merType": "L"
    }
    _form_data = urlencode(_form_data)
    with closing(requests.request("POST", BASE_QUERY_URL, headers=_request_headers, data=_form_data)) as response:
        try:
            price = response.text.split("<strong>")[1].split("</strong>")[0]  # 一句话将整个html解析出价格
        except:
            price = response.text.split("<h1>")[1].split("</h1>")[0] + "。"  # 当商品确实无报价，获得"对不起，该商品暂无回收商报价"的文本
            reason = response.text.split('<dd class="msg-body">')[1].split('</dd>')[0].replace('<br/>', '')
            price = price + reason
    #print(response.status_code)
    del _product_url
    del _spid
    del _radio_ids
    del _request_headers
    del _form_data
    return price

# 非常卡
def get_price_by_pool(_pool, _product_url, _radio_ids):
    _spid = _product_url.split("id=")[1].split("&")[0]  # _product_url与_spid本身存在一定关系
    _form_data = {
        "spId": _spid,
        "source": "",
        "typeId": "",
        "radioIds": _radio_ids,
        "multiIds": "",
        "areaId": "",
        "merType": "L"
    }
    _form_data = urlencode(_form_data)

    try:
        response = _pool.request("POST", BASE_QUERY_URL, headers=_pool.headers, body=_form_data)
    except:
        print("pool出异常了")
        traceback.print_exc()
    #with closing(requests.request("POST", BASE_QUERY_URL, headers=_request_headers, data=_form_data)) as response:
    price = response.data.split("<strong>")[1].split("</strong>")[0]  # 一句话将整个html解析出价格
    print(price)

    #print(response.status_code)
    del _product_url
    del _spid
    del _radio_ids
    #del _request_headers
    del _form_data
    return price

def get_price_by_session(_session, _product_url, _radio_ids):
    _spid = _product_url.split("id=")[1].split("&")[0]  # _product_url与_spid本身存在一定关系
    _request_headers = {
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8",
        "Accept-Encoding": "gzip, deflate",
        "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
        "Cache-Control": "max-age=0",
        "Connection": "close",
        "Content-Length": "438",
        "Content-Type": "application/x-www-form-urlencoded",
        "Host": "www.youdemai.net",
        "Origin": "http://www.youdemai.net",
        "Referer": _product_url,
        "Upgrade-Insecure-Requests": "1",
        "User-Agent": "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/69.0.3497.100 Safari/537.36"
    }

    _form_data = {
        "spId": _spid,
        "source": "",
        "typeId": "",
        "radioIds": _radio_ids,
        "multiIds": "",
        "areaId": "",
        "merType": "L"
    }
    _form_data = urlencode(_form_data)
    response = _session.post(BASE_QUERY_URL, headers=_request_headers, data=_form_data)
    #with closing(requests.request("POST", BASE_QUERY_URL, headers=_request_headers, data=_form_data)) as response:
    try:
        price = response.text.split("<strong>")[1].split("</strong>")[0]  # 一句话将整个html解析出价格
    except :
        price = response.text.split("<h1>")[1].split("</h1>")[0] + "。"  # 当商品确实无报价，获得"对不起，该商品暂无回收商报价"的文本
        reason = response.text.split('<dd class="msg-body">')[1].split('</dd>')[0].replace('<br/>', '')
        price = price + reason
    #print(response.status_code)
    del _product_url
    del _spid
    del _radio_ids
    del _request_headers
    del _form_data
    time.sleep(MIN_DELAY)
    return price


# Excel批量读、写json，dict，list等数据格式公共类
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


# python 实现N个数组的排列组合(笛卡尔积算法)
class Cartesian:
    # 初始化
    def __init__(self, datagroup):
        self.datagroup = datagroup
        # 二维数组从后往前下标值
        self.counterIndex = len(datagroup) - 1
        # 每次输出数组数值的下标值数组(初始化为0)
        self.counter = [0 for i in range(0, len(self.datagroup))]

    # 计算数组长度
    def countlength(self):
        i = 0
        length = 1
        while i < len(self.datagroup):
            length *= len(self.datagroup[i])
            i += 1
        return length

    # 递归处理输出下标
    def handle(self):
        # 定位输出下标数组开始从最后一位递增
        self.counter[self.counterIndex] += 1
        # 判断定位数组最后一位是否超过长度，超过长度，第一次最后一位已遍历结束
        if self.counter[self.counterIndex] >= len(self.datagroup[self.counterIndex]):
            # 重置末位下标
            self.counter[self.counterIndex] = 0
            # 标记counter中前一位
            self.counterIndex -= 1
            # 当标记位大于等于0，递归调用
            if self.counterIndex >= 0:
                self.handle()
            # 重置标记
            self.counterIndex = len(self.datagroup) - 1

    # 正向有序的(数组第一维度的顺序)排列组合输出
    def assemble(self):
        try:
            length = self.countlength()
            # print(length)  # debug only
            i = 0
            outlist = []
            while i < length:
                attrlist = []
                j = 0
                while (j < len(self.datagroup)):
                    attrlist.append(self.datagroup[j][self.counter[j]])
                    j += 1
                # print(attrlist)  # debug only
                outlist.append(attrlist)
                self.handle()
                i += self.density(length)  # 笛卡尔算法粗略化。原始值是 i +=1 ，这样算出来组合太多，内存消耗严重
            return outlist
        except:
            traceback.print_exc()

    def density(self, _length):  # 控制笛卡尔乘积的密度。当组合太多，电脑会卡死
        quotient, remainder = divmod(_length, AMOUNT)
        if quotient == 0:
            return 1
        else:
            return quotient


class JsonHelper:
    def write_json_to_txt(self, file_name, _json):
        with open(file_name, 'w+') as file_obj:
            '''写入json文件'''
            json.dump(_json, file_obj, indent=4, ensure_ascii=False)
            #print("写入json文件：", _json)

    def read_txt_as_json(self, file_name):
        with open(file_name) as file_obj:
            '''读取json文件'''
            _json = json.load(file_obj)  # 返回列表数据，也支持字典
            return _json



# 尝试用ProcessPoolExecutor代替ThreadPoolExecutor，绕过Global Interpretor Lock的限制
def get_formdata_and_price_by_url_batch(_brand, _product_name_list, _product_url_list):
    # MAX_WORKERS = os.cpu_count() 依赖当前运行环境的多核
    _max_workers = min(2, len(_product_url_list))
    with futures.ProcessPoolExecutor(max_workers=_max_workers) as executor:
        to_do = []
        count_future = 0  # 声明一个计数器，用来计算进程数，并方便打印
        #hp_not_done = ['惠普 CQ36', '惠普 dv2900', '惠普 dv6700', '惠普 6470b', '惠普 OMEN 17-an000TX', '惠普 4321s', '惠普 WASD 15-AX102TX', '惠普 4431s', '惠普 暗影精灵III代', '惠普 14-ar000', '惠普 WASD 15-AX101TX', '惠普 14-ar108TX', '惠普 14-r020tx', '惠普 15-ac636TX', '惠普 15g-ad002tx', '惠普 15g-br008TX', '惠普 15g-br009TX', '惠普 15g-bx001AX', '惠普 15g-bx002AX', '惠普 15g-bx004AX', '惠普 15q-aj001tx', '惠普 15q-aj006TX', '惠普 15q-aj103tx', '惠普 15q-aj107tx', '惠普 15q-aj108tx', '惠普 15q-aj111tx', '惠普 15q-aj122tx', '惠普 15q-bu103TX', '惠普 15q-bu104TX', '惠普 15q-by001AX', '惠普 17-ac002TX', '惠普 17-ac102TX', '惠普 17g-br001TX', '惠普 17g-br100TX', '惠普 17q-bu100TX', '惠普 240 G6', '惠普 245 G3', '惠普 245 G6', '惠普 250 G4', '惠普 250 G6', '惠普 346 G4', '惠普 348 G4', '惠普 4330s', '惠普 720 G2', '惠普 820 G2']
        #productname_not_done = ['华硕 X550', '华硕 N75', '华硕 ROG玩家国度系列 S7VS', '华硕 S7ZC']
        productname_not_done = get_incompleted_product_name(_brand, _product_name_list)  # 每次循环启动，就检查进度
        print("productname_not_done :" + str(len(productname_not_done)))
        # 用于创建future
        for i in range(len(_product_name_list)):
            if _product_name_list[i] not in productname_not_done: continue
            #if product_name_list[i] != "戴尔 Latitude E6430s" : continue
            #if i < 566: continue
            # submit方法排定可调用对象的执行时间后返回一个future，表示这个待执行的操作
            # submit第一个参数是 获取格式数据和价格的方法get_formdata_and_price_by_url
            # 第二个参数是 submit第一个参数(方法名)的参数的 一个子元素， 即每次循环遍历出来的一个
            _future = executor.submit(get_formdata_and_price_by_url, _brand, _product_url_list[i])
            to_do.append(_future)
            print("++++++++++在执行第  " + str(count_future) + "  个循环++++++++++")
            #time.sleep(1)
            count_future += 1

        results = []
        # 用于获取future的结果
        # as_completed 接收一个future列表，返回值是一个迭代器，在运行结束后产出future

        for future in futures.as_completed(to_do):
            _result = future.result()
            print("收到结果")
            results.append(_result)
    return len(results)

def main(get_formdata_and_price_by_url_batch):  # <10>
    t0 = time.time()
    count = get_formdata_and_price_by_url_batch(out_brand, out_product_name_list, out_product_url_list)
    elapsed = time.time() - t0
    msg = '\n{} flags downloaded in {:.2f}s'
    print(msg.format(count, elapsed))


if __name__ == '__main__':

    print(type(BRAND_DICT))
    print(MAX_WORKERS)

    notebook_url = "http://www.youdemai.net/products/product/brands?c=5&type=L&source="
    brand_dict = get_brand_dict(notebook_url)
    out_brand = "asus"


    out_product_url_list, spid_list = get_producturls_by_url(out_brand)
    print(len(out_product_url_list))

    # 产品名列表，总表
    out_product_name_list = get_productnames_from_producturls(out_product_url_list)

    # 直接运行
    #for i in range(len(product_url_list)):
    #    #if i < 10: continue  # 总共就16台，  i=6
    #    print("开始执行main方法里的循环，第" +str(i+1) +"次")
    #    get_formdata_and_price_by_url(product_url_list[i])
    #    time.sleep(3)

    # 多进程
    main(get_formdata_and_price_by_url_batch)  # 4进程 apple 200条 767秒  16进程 apple 500条 520秒
