# -*- coding: utf-8 -*-

if __name__ == "__main__":
    reader = open(r'searchTest.txt', encoding='utf-8', errors='ignore')  # 把searchTest.txt换成你的文件的绝对路径
    data = reader.readline()  # 读取一行数据
    odercode_list = []  # 订单号集合
    method_list = []  # 方法集合
    timestamp_list = []  # 时间戳集合
    while data != '':  # 循环读取数据行
        data = reader.readline()
        if "<orderCode>" in data:
            tmp = str.split(data, "<orderCode>")[1].strip("\n")
            odercode = str.split(tmp, "</orderCode>")[0].strip("\n")
            odercode_list.append(odercode)

        if "rtx.controller.RtxController - 方法" in data:
            method = str.split(data, "：")[1].strip("\n")  # 把方法拿出来
            method_list.append(method)

        if "时间戳" in data:   # 如果一行数据包含时间戳
            time = str.split(data, "：")[1].strip("\n")   # 拆分一下，把时间取出来
            timestamp_list.append(time)   # 添加到集合里

    reader.close()
    print(odercode_list)
    print(method_list)
    print(timestamp_list)
