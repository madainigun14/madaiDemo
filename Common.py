import unittest

# 省会字典
administrative_division_dict ={
    "北京市",
    "上海市",
    "天津市",
    "重庆市",
    "广州市",
    "呼和浩特市",
    "乌鲁木齐市",
    "拉萨市",
    "银川市",
    "宁波市",
    "长沙市",
    "苏州市",
    "杭州市",
    "昆明市",
    "成都市",
    "兰州市",
    "南京市",
    "武汉市",
    "贵阳市",
    "南昌市",
    "深圳市",
    "南宁市",
    "西宁市",
    "西安市",
    "合肥市",
    "福州市",
    "海口市",
    "石家庄市",
    "太原市",
    "长春市",
    "哈尔滨市",
    "沈阳市",
    "香港特别行政区",
    "澳门特别行政区"
}




class SequentialTestLoader(unittest.TestLoader):
    def getTestCaseNames(self, testCaseClass):
        test_names = super().getTestCaseNames(testCaseClass)
        testcase_methods = list(testCaseClass.__dict__.keys())
        test_names.sort(key=testcase_methods.index)
        return test_names