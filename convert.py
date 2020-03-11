# coding:utf-8
import os
import sys
reload(sys)
sys.setdefaultencoding("utf-8")

from excelConfig import ExcelConfig 
from enums import Significance
from enums import ExecMode

class Convert():
    def __init__(self, ExcelFileName, SheetName):
        self.excelFile = ExcelFileName + '.xlsx'
        self.excelSheet = SheetName
        self.temp = ExcelConfig(self.excelFile)
        self.dic_testlink = {}
        self.row_flag = 3
        self.testsuite = self.temp.getCell(self.excelSheet, 2, 1)
        self.dic_testlink[self.testsuite] = {"node_order": "13", "details": "", "testcase": []}
        self.content = ""
        self.content_list = []

    def xlsx_to_dic(self, SheetName):
        
        # node_order = int(self.temp.getCell(self.excelSheet,self.row_flag - 1,9))
        # externalId = int(self.temp.getCell(self.excelSheet,self.row_flag - 1,10))
        while True:
            # print 'loop1'
            # list_testcase = dic_testlink[testsuite].["testcase"]

            testcase = {"name": "", "node_order": "1000", "externalid": "", "version": "1", "summary": "",
                        "preconditions": "",  "execution_mode": "1", "significance": "2","status":"1", "steps": [], "keywords": "1.0"}
            testcase["name"] = self.temp.getCell(self.excelSheet, self.row_flag, 1)
            # testcase["node_order"] = node_order
            # testcase["externalid"] = externalId
            testcase["summary"] = self.temp.getCell(self.excelSheet, self.row_flag, 4)
            testcase["preconditions"] = self.temp.getCell(self.excelSheet, self.row_flag, 5) 
            significance = self.temp.getCell(self.excelSheet,self.row_flag,6)
            execution_mode = self.temp.getCell(self.excelSheet, self.row_flag, 3)

            # type1 = type(significance)                    unicode：默认使用utf-8将“重要性”中文值解码（decode）成了unicode
            # type2 = type(execution_mode.encode('utf-8'))  str：手动编码（encode），将unicode转换成str
            # type3 = type(ExecMode.auto)                   Enum：未进行任何处理，auto对象默认还是枚举类型
            # type4 = type(ExecMode.auto.value)             str：枚举类型的值，该枚举对象的值类型是str
            # 这里对中文字符进行比较，需要转换成相同（编码）类型；如str，或者unicode
            # execution_mode解码之前是str，所以ExecMode.auto枚举需要取value值（str类型），再解码成unicode
            if execution_mode == ExecMode.auto.value.decode('utf-8'):
                testcase["execution_mode"] = 2

            if significance is None:
                raise Exception("significance is required!")
            elif significance.strip() == Significance.high.value.decode('utf-8'):
                testcase["significance"] = 3
            elif significance.strip() == Significance.medium.value.decode('utf-8'):
                testcase["significance"] = 2
            else :
                testcase["significance"] = 1
            # print self.temp.getCell('Sheet1',self.row_flag,3)
            step_number = 1 
            testcase["keywords"] = self.temp.getCell(self.excelSheet, self.row_flag, 2)
            if testcase["keywords"] is not None:
                testcase["keywords"].strip()
            else:
                raise Exception("Keywords is required!")
            # node_order += 1
            # externalId += 1
            # print testcase["keywords"]
            while True:
                # print 'loop2'
                step = {"step_number": "", "actions": "", "expectedresults": "", "execution_mode": "1"}
                step["step_number"] = step_number
                step["actions"] = self.temp.getCell(self.excelSheet, self.row_flag, 7)
                step["expectedresults"] = self.temp.getCell(self.excelSheet, self.row_flag, 8)
                if execution_mode == ExecMode.auto.value.decode('utf-8'):
                    step["execution_mode"] = 2
                testcase["steps"].append(step)
                step_number += 1
                self.row_flag += 1
                if self.temp.getCell(self.excelSheet, self.row_flag, 1) is not None or self.temp.getCell(self.excelSheet, self.row_flag, 7) is None:
                    break
            # print testcase

            self.dic_testlink[self.testsuite]["testcase"].append(testcase)
            # print self.row_flag
            if self.temp.getCell(self.excelSheet, self.row_flag, 7) is None and self.temp.getCell(self.excelSheet, self.row_flag + 1, 7) is None:
                break
        self.temp.close()
        # print self.dic_testlink

    def content_to_xml(self, key, value=None):
        if key == 'step_number' or key == 'execution_mode' or key == 'node_order' or key == 'externalid' or key == 'version' or key == 'significance':
            return "<" + str(key) + "><![CDATA[" + str(value) + "]]></" + str(key) + ">"
        elif key == 'actions' or key == 'expectedresults' or key == 'summary' or key == 'preconditions':
            return "<" + str(key) + "><![CDATA[<p> " + str(value) + "</p> ]]></" + str(key) + ">"
        elif key == 'keywords':
            return '<keywords><keyword name="' + str(value) + '"><notes><![CDATA[]]></notes></keyword></keywords>'
        elif key == 'name':
            return '<testcase name="' + str(value) + '">'
        elif key == 'status':
            return "<status>" + str(value) + "</status>"
        else:
            return '##########'

    def dic_to_xml(self, ExcelFileName, SheetName):
        testcase_list = self.dic_testlink[self.testsuite]["testcase"]
        for testcase in testcase_list:
            for step in testcase["steps"]:
                self.content += "<step>"
                self.content += self.content_to_xml("step_number", step["step_number"])
                self.content += self.content_to_xml("actions", step["actions"])
                self.content += self.content_to_xml("expectedresults", step["expectedresults"])
                self.content += self.content_to_xml("execution_mode", step["execution_mode"])
                self.content += "</step>"
            self.content = "<steps>" + self.content + "</steps>"
            self.content = self.content_to_xml("status",testcase["status"]) + self.content
            self.content = self.content_to_xml("significance", testcase["significance"]) + self.content
            self.content = self.content_to_xml("execution_mode", testcase["execution_mode"]) + self.content
            self.content = self.content_to_xml("preconditions", testcase["preconditions"]) + self.content
            self.content = self.content_to_xml("summary", testcase["summary"]) + self.content
            self.content = self.content_to_xml("version", testcase["version"]) + self.content
            #self.content = self.content_to_xml("externalid", testcase["externalid"]) + self.content
            #self.content = self.content_to_xml("node_order", testcase["node_order"]) + self.content   
            self.content = self.content + self.content_to_xml("keywords", testcase["keywords"])
            self.content = self.content_to_xml("name", testcase["name"]) + self.content
            self.content = self.content + "</testcase>"
            self.content_list.append(self.content)
            self.content = ""
        self.content = "".join(self.content_list)

        # 根据excel数据源确定是否需要生成外层用例集名称
        if self.testsuite is not None:
            self.content = self.content_to_xml("keywords",testcase["keywords"]) + self.content
            self.content = '<testsuite name="' + self.testsuite + '">' + self.content + "</testsuite>"
        else:
            self.content = "<testcases>" + self.content + "</testcases>"

        self.content = '<?xml version="1.0" encoding="UTF-8"?>' + self.content
        self.write_to_file(ExcelFileName, SheetName)

    def write_to_file(self, ExcelFileName, SheetName):
        xmlFileName = 'output\\' + ExcelFileName + '_' + SheetName + '.xml'
        cp = open(xmlFileName, "w")
        cp.write(self.content)
        cp.close()

if __name__ == "__main__":
    # res = os.system('pip install -r .\\dependency\\requirements.txt')   
    # print res
    fileName = raw_input('Enter excel name:').strip()
    sheetName = raw_input('Enter sheet name:').strip()
    sheetList = sheetName.split(" ")
    for sheetName in sheetList:
        test = Convert(fileName, sheetName)
        test.xlsx_to_dic(sheetName)
        test.dic_to_xml(fileName, sheetName)
    print "Convert successfully!"
    os.system('pause')
