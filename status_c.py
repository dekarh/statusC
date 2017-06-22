# -*- coding: utf-8 -*-

import sys
import time
import openpyxl
# from openpyxl import Workbook
import csv
from lib import IN_NAME, IN_SNILS, IN_STAT, OUT_NAME, OUT_STAT

workbooks =  []
sheets = []
for i, xlsx_file in enumerate(sys.argv):                              # Загружаем все xlsx файлы
    if i == 0:
        continue
    workbooks.append(openpyxl.load_workbook(filename=xlsx_file, read_only=True))
    sheets.append(workbooks[i-1][workbooks[i-1].sheetnames[0]])
#    for j, row in enumerate(sheets[i-1].rows):
#        for k, cell in enumerate(row):
#            g=0

sheets_keys = []
for i, sheet in enumerate(sheets):                                    # Маркируем нужные столбцы
    keys = {}
    for j, row in enumerate(sheet.rows):
        if j > 0:
            break
        for k, cell in enumerate(row):
            if cell.value in IN_SNILS:
                keys[IN_SNILS[0]] = k
        if len(keys) > 0:
            for k, cell in enumerate(row):
                for name in IN_NAME:
                    if cell.value != None:
                        if cell.value == name:
                            keys[name] = k
        else:
            print('В файле ' + sys.argv[i+1] + 'отсутствует колонка со СНИЛС')
            time.sleep(3)
            sys.exit()
    sheets_keys.append(keys)

for j, row in enumerate(sheets[0].rows):                     # Загружаем все входные данные в одну строку
    if j == 0:
        continue
    big_row = {}
    for k, sheet_key in enumerate(sheets_keys[0]):
        big_row[sheet_key] = row[sheets_keys[0][sheet_key]].value
    for i, sheet in enumerate(sheets):
        if i == 0:
            continue
        if str(type(big_row[IN_SNILS[0]])) == "<class 'str'>":
            if big_row[IN_SNILS[0]].strip() != '':
                for jj, row in enumerate(sheets[i].rows):
                    if str(type(row[sheets_keys[i][IN_SNILS[0]]].value)) == "<class 'str'>":
                        if row[sheets_keys[i][IN_SNILS[0]]].value.strip() == big_row[IN_SNILS[0]].strip():
                            for k, sheet_key in enumerate(sheets_keys[i]):
                                big_row[sheet_key] = row[sheets_keys[i][sheet_key]].value
                            break
    g = 0


stat_our2csv = [{'Имя':'Он они он','Возраст':25,'Вес':200},
         {'Имя':'Я я я','Возраст':31,'Вес':180}]
stat_fond2csv = [{'Имя':'Он они он','Возраст':25,'Вес':200},
         {'Имя':'Я я я','Возраст':31,'Вес':180}]
with open('stat_our.csv', 'w', encoding='cp1251') as output_file:
    dict_writer = csv.DictWriter(output_file, stat_our2csv[0].keys(), quoting=csv.QUOTE_NONNUMERIC)
    dict_writer.writeheader()
    dict_writer.writerows(stat_our2csv)
output_file.close()
with open('stat_fond.csv', 'w', encoding='cp1251') as output_file:
    dict_writer = csv.DictWriter(output_file, stat_fond2csv[0].keys(), quoting=csv.QUOTE_NONNUMERIC)
    dict_writer.writeheader()
    dict_writer.writerows(stat_fond2csv)
output_file.close()
