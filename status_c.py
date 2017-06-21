# -*- coding: utf-8 -*-

import sys
import openpyxl
# from openpyxl import Workbook
import csv

workbook = openpyxl.load_workbook(filename=sys.argv[1], read_only=True)
sheet = workbook.active
for i, row in enumerate(sheet.rows):
        for cell in row:
            print(cell.value)

# open('input.txt', 'r',  encoding='cp1251')

