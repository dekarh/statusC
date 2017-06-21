# -*- coding: utf-8 -*-

import sys
import openpyxl
# from openpyxl import Workbook
import csv

workbook = openpyxl.load_workbook(filename=sys.argv[1], read_only=True)
sheet = workbook.active
g = 0
for i, row in enumerate(sheet.rows):
        for j, cell in enumerate(row):
            g=0

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
