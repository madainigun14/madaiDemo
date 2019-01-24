import xlrd
import json
import xlwings
import win32con
import win32ui
import Util

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

            dictdata = dict(zip(_titlelist, rowvalues))     # 将title列作为Key，将其余的列作为Value，循环打包成字典
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
            rows = currentsheet.nrows    # 当前的Sheet里有nrows行

            # 获取第一行数据，作为标题数据，也就是dict里的key
            _titlelist = currentsheet.row_values(0)

            # 定义一个list存放所有非第一行数据
            # 非第一行循环添加为value
            dictlist = []
            for row in range(1, rows):  # 从第2行开始
                rowvalues = currentsheet.row_values(row)
                # print("每行的数据是：")
                # print(rowValues) #debug only，result：PASS

                dictdata = dict(zip(_titlelist, rowvalues))     # 将title列作为Key，将其余的列作为Value，循环打包成字典
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


if __name__ == "__main__":
    # 打开文件选择框
    fspec = "Excel文件 (*.xlsx)|*.xlsx|Excel文件 (*.xls)|*.xls"
    flags = win32con.OFN_OVERWRITEPROMPT | win32con.OFN_FILEMUSTEXIST

    dialog = win32ui.CreateFileDialog(1, None, None, flags, fspec)  # 1表示打开文件对话框
    dialog.SetOFNInitialDir("C:\\")  # 初始位置
    dialog.SetOFNTitle("请选择一个合法的Excel文件")  # 对话框的标题
    dialog.DoModal()
    filepath = dialog.GetPathName()

    # 初始化Excel对象，用来接收传过来的Json
    # excelobj = Excel("F:\新闻管理系统\Test Data\新闻管理(多个sheet).xlsx")
    excelobj = Util.Excel(filepath)
    data = excelobj.readsheets()
    print("sheetdata的值是 \r\n" + str(data))  # debug only

    # 另存为对话框
    dialog = win32ui.CreateFileDialog(0, None, None, flags, fspec)  # 0表示另存为对话框
    dialog.SetOFNInitialDir("C:\\")
    dialog.DoModal()
    _path = dialog.GetPathName()

    excel = Util.Excel(_path)
    excel.writesheets(data)

