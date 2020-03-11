# coding=utf-8
from xml.etree import ElementTree
from win32com.client import Dispatch
import win32com.client
import os
import sys
reload(sys)
sys.setdefaultencoding("utf-8")

class ExcelConfig:
    def __init__(self, filename=None):
        self.xlApp = win32com.client.Dispatch('Excel.Application')

        #path =  os.getcwd()    os.getcwd() 方法用于返回当前工作目录，所以在其他文件 import 了我们的包之后，os.getcwd() 返回的是：当前正运行的 python 文件目录
        #path1 = __file__
        #path2 = os.path.realpath(__file__)
        #path3 = os.path.dirname(__file__)
        #path4 = os.path.dirname(os.path.realpath(__file__))
        if filename:
            self.filename = os.getcwd() + "\\sourcefile\\" + filename
            # self.xlApp.Visible=True
            self.xlBook = self.xlApp.Workbooks.Open(self.filename)
        else:
            # self.xlApp.Visible=True
            self.xlBook = self.xlApp.Workbooks.Add()
            self.filename = ''

    def save(self, newfilename=None):
        if newfilename:
            self.filename = os.getcwd() + "\\" + newfilename
            # if os.path.exists(self.filename):
            # os.remove(self.filename)
            self.xlBook.SaveAs(self.filename)
        else:
            self.xlBook.Save()

    def close(self):
        self.xlBook.Close(SaveChanges=0)
        self.xlApp.Quit()

    def getCell(self, sheet, row, col):
        sht = self.xlBook.Worksheets(sheet)
        return sht.Cells(row, col).Value

    def setCell(self, sheet, row, col, value):
        sht = self.xlBook.Worksheets(sheet)
        sht.Cells(row, col).Value = value
        # 设置居中
        sht.Cells(row, col).HorizontalAlignment = 3
        sht.Rows(row).WrapText = True

    def mergeCells(self, sheet, row1, col1, row2, col2):
        start_coloum = int(dic_config["start_coloum"])
        # 如果这列不存在就不合并单元格
        if col2 != start_coloum - 1:
            sht = self.xlBook.Worksheets(sheet)
            sht.Range(sht.Cells(row1, col1), sht.Cells(row2, col2)).Merge()
            # else:
            # print 'Merge cells coloum %s failed!' %col2

    def setBorder(self, sheet, row, col):
        sht = self.xlBook.Worksheets(sheet)
        sht.Cells(row, col).Borders.LineStyle = 1

    def set_col_width(self, sheet, start, end, length):
        start += 96
        end += 96
        msg = chr(start) + ":" + chr(end)
        # print msg
        sht = self.xlBook.Worksheets(sheet)
        sht.Columns(msg.upper()).ColumnWidth = length