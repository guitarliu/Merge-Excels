# coding: utf-8

import os, re
from xlrd import *
from xlwt import *
from xlutils.copy import copy
from win32com import client
from openpyxl import *


class Book_Info(object):
    def __init__(self,b1,b2):
        self.b1_primary = b1
        self.b2_primary = b2
    def get_rc(self, bookname):
        r=bookname.sheets()[0].nrows
        c=bookname.sheets()[0].ncols
        return r,c
    def get_cvalues(self, bookname, r_num):
        return bookname.sheets()[0].row_values(r_num)
    def convert_xls(self, excel_name):
        excel = client.gencache.EnsureDispatch("Excel.Application")
        wb = excel.Workbooks.Open(os.getcwd() + "/" + excel_name)    # must be absolute path
        wb.SaveAs(os.getcwd() + "/" + excel_name.split("xlsx")[0] + "xls", FileFormat = 56)  # need to add the absolute path
        wb.Close()
        excel.Application.Quit()
    def convert_xlsx(self, excel_name):
        excel = client.gencache.EnsureDispatch("Excel.Application")
        wb = excel.Workbooks.Open(os.getcwd() + "/" + excel_name)    # must be absolute path
        wb.SaveAs(os.getcwd() + "/" + excel_name + "x", FileFormat = 51)  # need to add the absolute path
        wb.Close()
        excel.Application.Quit()
    def compare_contents(self):
        '''convert to xls and then compare them'''
        self.convert_xls(self.b1_primary)
        self.convert_xls(self.b2_primary)
        b1_xls=open_workbook(os.getcwd() + "/" + self.b1_primary.split("xlsx")[0] + "xls", formatting_info=True, on_demand=False)
        b2_xls=open_workbook(os.getcwd() + "/" + self.b2_primary.split("xlsx")[0] + "xls", formatting_info=True, on_demand=False)
        wb1=copy(b1_xls)
        for s1 in range(self.get_rc(b1_xls)[0]):
            for s2 in range(self.get_rc(b2_xls)[0]):
                if self.get_cvalues(b1_xls, s1)[2:5]  == self.get_cvalues(b2_xls, s2)[2:5]:
                    wb = load_workbook(os.getcwd() + "/" + self.b2_primary)
                    ws = wb.active
                    for row in ws.rows:
                        if [row_content.value for row_content in row][2:5] == self.get_cvalues(b2_xls, s2)[2:5]:
                            ws.delete_rows(int(re.findall(r"\d+", str(row[0]))[-1]))
                            wb.save(os.getcwd() + "/" + self.b2_primary)
                    for num in range(6, 13):
                        if b1_xls.sheets()[0].cell_value(s1, num) == '':
                            result = b2_xls.sheets()[0].cell_value(s2, num)
                        elif b2_xls.sheets()[0].cell_value(s2, num) == '':
                            result = b1_xls.sheets()[0].cell_value(s1, num)
                        elif b1_xls.sheets()[0].cell_value(s1, num) != '' and b2_xls.sheets()[0].cell_value(s2, num) != '' :
                            result = b1_xls.sheets()[0].cell_value(s1, num) + b2_xls.sheets()[0].cell_value(s2, num)
                        wb1.get_sheet(0).write(s1, num, result)
        wb1.save("result.xls")

book_1=Book_Info("1.xlsx", "2.xlsx")
book_1.compare_contents()
#book_result = Book_Info("result.xls", "2.xlsx")
#book_result.convert_xls("2.xlsx")

# write contents of 2.xls into result.xls
#middle_xls = open_workbook(os.getcwd() + "/" + "2.xls")
#nrows, ncols = book_result.get_rc(middle_xls)[0], book_result.get_rc(middle_xls)[1]
#for row in range(nrows):
#   for col in range(ncols):
#      workbook = open_workbook(os.getcwd() + "/" + "result.xls")
#        workbooknew = copy(workbook)
#        ws = workbooknew.get_sheet(0)
#        row_start = len(ws.get_rows())
#        contents = middle_xls.sheet_by_index(0).cell(row, col).value
#        ws.write(row_start, col, contents)
#        workbooknew.save("final.xls")





