# coding: utf-8

import os, re, numpy
from xlrd import *
from xlwt import *
from xlutils.copy import copy
from win32com import client
from openpyxl import *


# extract the same contents and write them into a middle xlsx file
# extract the different contents of the second xlsx file and write them into a middle xlsx file
wb1 = load_workbook(os.getcwd() + "/" + "1.xlsx")
wb2 = load_workbook(os.getcwd() + "/" + "2.xlsx")
wb_final = Workbook()
ws1 = wb1.active
ws2 = wb2.active
ws_final = wb_final.active

for row1 in ws1.rows:
    for row2 in ws2.rows:
        if [row1_content.value for row1_content in row1][2:5] == \
                [row2_content.value for row2_content in row2][2:5]:
            contents = list(numpy.array([row1_content.value for row1_content in row1][6:13]) +
                  numpy.array([row2_content.value for row2_content in row2][6:13]))
            final_contents = [row1_content.value for row1_content in row1][:6] + contents + \
                             [row1_content.value for row1_content in row1][13::]
            ws_final.append(final_contents)
            wb_final.save(os.getcwd() + "/" + "middle.xlsx")
            ws2.delete_rows(int(re.findall(r"\d+", str(row2[0]))[-1]))
            wb2.save(os.getcwd() + "/" + "2" + "_middle.xlsx")

# extract the different contents of the first xlsx file and write them into a middle xlsx file
wb_middle = load_workbook(os.getcwd() + "/" + "middle.xlsx")
wb1_a = load_workbook(os.getcwd() + "/" + "1.xlsx")
ws_middle = wb_middle.active
ws1_a = wb1_a.active

for row1 in ws_middle.rows:
    for row2 in ws1_a.rows:
        if [row1_content.value for row1_content in row1][2:5] == \
                [row2_content.value for row2_content in row2][2:5]:
            ws1_a.delete_rows(int(re.findall(r"\d+", str(row2[0]))[-1]))
            wb1_a.save(os.getcwd() + "/" + "1" + "_middle.xlsx")

# write two middle xlsx files into the middle file
wb_final = load_workbook(os.getcwd() + "/" + "middle.xlsx")
wb1_middle = load_workbook(os.getcwd() + "/" + "1" + "_middle.xlsx")
wb2_middle =load_workbook(os.getcwd() + "/" + "2" + "_middle.xlsx")
ws_final = wb_final.active
ws1_middle = wb1_middle.active
ws2_middle = wb2_middle.active
for row1 in ws1_middle.rows:
    ws_final.append([row1_content.value for row1_content in row1])
    wb_final.save(os.getcwd() + "/" + "final.xlsx")
for row2 in ws2_middle.rows:
    ws_final.append([row2_content.value for row2_content in row2])
    wb_final.save(os.getcwd() + "/" + "final.xlsx")