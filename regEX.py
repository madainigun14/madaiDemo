# -*- coding: utf-8 -*-
import tempfile
import win32api
import win32ui
import win32con
import win32print
import re


if __name__ == "__main__":
	key = r"l㐓驲ala阿萨德刚阿萨德刚阿萨德刚撒是la是<hT是ml>hell是是o</Htm是l>hei是he是萨的嘎是撒as噶十多个ihei是"
	p1 = r"[\u4e00-\u9fa5]"
	pattern1 = re.compile(p1)
	print(pattern1.findall(key))