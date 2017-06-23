# -*- coding: utf-8 -*-

import sys
import time
import openpyxl
# from openpyxl import Workbook
import csv
from lib import IN_NAME, IN_SNILS, IN_STAT_OUR, IN_STAT_FOND , OUT_NAME, OUT_STAT

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
        for k, cell in enumerate(row):                                # Проверяем, чтобы был СНИЛС
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

our_statuses = []
for j, row in enumerate(sheets[0].rows):                     # Загружаем все входные данные в одну строку
    our_status = {}
    if j == 0:
        continue
    big_row = {}
    for k, sheet_key in enumerate(sheets_keys[0]):
        if row[sheets_keys[0][sheet_key]].value != None \
                and str(row[sheets_keys[0][sheet_key]].value).strip() != '':
              # and str(row[sheets_keys[0][sheet_key]].value).strip() != '—'
            big_row[sheet_key] = str(row[sheets_keys[0][sheet_key]].value)
    for i, sheet in enumerate(sheets):
        if i == 0:
            continue
        if str(type(big_row[IN_SNILS[0]])) == "<class 'str'>":
            if big_row[IN_SNILS[0]].strip() != '':
                for row in sheets[i].rows:
                    if str(type(row[sheets_keys[i][IN_SNILS[0]]].value)) == "<class 'str'>":
                        if row[sheets_keys[i][IN_SNILS[0]]].value.strip() == big_row[IN_SNILS[0]].strip():
                            for k, sheet_key in enumerate(sheets_keys[i]):
                                tek = row[sheets_keys[i][sheet_key]].value
                                if tek != None and str(tek).strip() != '':                # and str(tek).strip() != '—'
                                    if str(tek)[:12].lower() == 'одобрено инф':
                                        big_row[sheet_key] = OUT_STAT['Фонд - Статус КоллЦентра'][13]
                                    elif str(tek)[:15].lower() == 'акцепт прозвона':
                                        big_row[sheet_key] = OUT_STAT['Фонд - Статус КоллЦентра'][12]
                                    else:
                                        big_row[sheet_key] = str(row[sheets_keys[i][sheet_key]].value)
                            break
    for i, name in enumerate(OUT_NAME): # Заполняем список-
        if i == 0:
            our_status[name] = big_row[name]
        else:
            our_status[name] = IN_STAT_OUR[name].index(big_row[name])
    for i, name in enumerate(OUT_NAME):
        if i == 0:
            continue
# Вручную, Бумага принята только если в обоих полях Исправили или Наличие бумаги
        elif name == 'СтатусБумажногоНосителяПоДоговору' or name == 'СтатусБумажногоНосителяПоЗаявлению':
            tek1 = big_row['СтатусБумажногоНосителяПоДоговору']
            tek2 = big_row['СтатусБумажногоНосителяПоЗаявлению']
            if (tek1 == 'Исправили' or tek1 == 'Наличие бумаги') and (tek2 == 'Исправили' or tek2 == 'Наличие бумаги'):
                our_status[name] = OUT_STAT['Фонд - Статус бумаги'].index('Бумага принята')
            else:
                our_status[name] = OUT_STAT['Фонд - Статус бумаги'].index('Ошибка')
        else:
            our_status[name] = IN_STAT_FOND[name][big_row[name]]
    g = 0






        # our_statuses = [{'Имя':'Он они он','Возраст':25,'Вес':200}, {'Имя':'Я я я','Возраст':31,'Вес':180}]
with open('statuses.csv', 'w', encoding='cp1251') as output_file:
    dict_writer = csv.DictWriter(output_file, our_statuses[0].keys(), quoting=csv.QUOTE_NONNUMERIC)
    dict_writer.writeheader()
    dict_writer.writerows(our_statuses)
output_file.close()
