# -*- coding: utf-8 -*-
import tempfile
import win32api
import win32ui
import win32con
import win32print


def printpdf():
	filename = tempfile.mktemp(".txt")
	open(filename, "w").write("This is a test")
	win32api.ShellExecute(
		0,
		"print",
		filename,
		#
		# If this is None, the default printer will
		# be used anyway.
		#
		'/d:"%s"' % win32print.GetDefaultPrinter(),
		".",
		0
	)


def print2printer():
	"""
	Device Context绘图的方式进行打印
	:return:
	"""
	inch = 1440
	hdc = win32ui.CreateDC()
	hdc.CreatePrinterDC(win32print.GetDefaultPrinter())
	hdc.StartDoc("Test doc")
	hdc.StartPage()
	hdc.SetMapMode(win32con.MM_TWIPS)
	hdc.DrawText("Hello, World!", (0, inch * -1, inch * 8, inch * -2), win32con.DT_CENTER)
	hdc.EndPage()
	hdc.EndDoc()


if __name__ == "__main__":
	# print2printer()
	printpdf()
