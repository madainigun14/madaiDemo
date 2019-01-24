#  -*- coding: utf-8 -*-
import unittest,  time,  re,  os
import xlrd,  xlwt,  xlutils,  openpyxl
import json
import requests
from requests import adapters  #  requests.adapters.DEFAULT_RETRIES = 5
import traceback
import xlwings
from urllib.parse import urlencode
import urllib,  urllib3
from contextlib import closing
import sys
from lxml import etree
import gc,  objgraph  #  监控内存
import threading
import socket
import re
import os
import http.client

if __name__ == '__main__':


    conn = http.client.HTTPConnection("www.youdemai.net")

    payload = ""

    headers = {
        'accept': "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8",
        'accept-encoding': "gzip, deflate",
        'accept-language': "zh-CN,zh;q=0.9,en;q=0.8",
        'cache-control': "no-cache",
        'content-type': "application/x-www-form-urlencoded",
        'host': "www.youdemai.net",
        'origin': "http//www.youdemai.net",
        'upgrade-insecure-requests': "1",
        'user-agent': "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/69.0.3497.100 Safari/537.36",
        'postman-token': "e7b6cb63-1fa1-2e5f-d970-d06c998b7d9d"
    }
    formdata = "spId=689AF40E132447D9E050840A65081DF1&source=&typeId=&radioIds=6D17F1261705B11CE050840A650801D2%236D17F1261702B11CE050840A650801D2%236D17F12616FFB11CE050840A650801D2%236D17F12616F9B11CE050840A650801D2%236D17F12616F7B11CE050840A650801D2%236D17F12616F2B11CE050840A650801D2%23732E6CFEA86707ABE050840A65084197%23732E6BECD7226FB4E050840A650840CD%23732E68445F2340DDE050840A65083E2F%23732E66FCEC4A4407E050840A65083D79&multiIds=&areaId=&merType=L"
    _form_data = {
        "spId": "105CD2BD7E91416DE050007F01002965",
        "source": "",
        "typeId": "",
        "radioIds": "6D17D40E754A0AF9E050840A65085343#6D17D40E75470AF9E050840A65085343#6D17D40E75440AF9E050840A65085343#6D17D40E753E0AF9E050840A65085343#6D17D40E753A0AF9E050840A65085343#6D17D40E75370AF9E050840A65085343#7D4915FA31E19363E050840A6508778A#7D4913652FAE7564E050840A650873DE#7D49127E435C3F2FE050840A650871A1#7D4910E6F3F8D55FE050840A65086E9F",
        "multiIds": "",
        "areaId": "",
        "merType": "L"
    }

    conn.request("POST", "/order/order/inquiry", payload, str(headers)+str(_form_data))

    res = conn.getresponse()
    data = res.read()

    print(data)


