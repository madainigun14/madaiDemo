# -*- coding: utf-8 -*-
import requests
import json


# 获取token
def gettoken(url, data):
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
def getresponsedata(url, data):
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
response = getresponsedata(url, data)
token = gettoken(url, data)

print(response)
print(token)