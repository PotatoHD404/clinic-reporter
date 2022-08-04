import re

import pdfplumber
import camelot
# creating a pdf file object
# import ctypes
# from ctypes.util import find_library
# print(find_library("".join(("gsdll", str(ctypes.sizeof(ctypes.c_voidp) * 8), ".dll"))))
# 0 - №
# 1 - Время приема
# 2 - Пациент
# 3 - Медицинский работникк
# 4 Назначенные исследования
# 5 - Код диагноза
# 6 - Код услуги (кратность)
import unicodedata
import glob
import pandas as pd

months = {'Января': '01',
          'Февраля': '02',
          'Марта': '03',
          'Апреля': '04',
          'Мая': '05',
          'Июня': '06',
          'Июля': '07',
          'Августа': '08',
          'Сентября': '09',
          'Октября': '10',
          'Ноября': '11',
          'Декабря': '12'}
data = {}
regex = r"([0-9]{5}) \(([0-9]+)\)"
for path in glob.glob("data/*.pdf"):
    with pdfplumber.open(path) as file:
        date = file.pages[0].extract_text().split('Дата приема: ')[1].split('\n')[0]
        if date not in data:
            data[date] = {}
        tables = camelot.read_pdf(path, pages='all')
        for table in tables:
            for el in table.df.values[1:]:
                data[date][el[1]] = [{'type': int(i[0]), 'count': int(i[1])} for i in
                                     re.findall(regex, unicodedata.normalize("NFKD", el[6]))]
# print(data)
result = []
for key, value in sorted(data.items()):
    date = months[key.split(' ')[1]] + '.' + key.split(' ')[0]
    result.append(
        {'Дата': date, 'Чел': 0, 40001: 0, 40004: 0, 40032: 0, 40051: 0, 40054: 0, 40033: 0, 40036: 0, 40052: 0, 40003: 0,
         40055: 0, 40035: 0, 40021: 0, 40101: 0, 40091: 0, 40093: 0, 40095: 0, 40083: 0, 40072: 0,
         'Кратность человек': 0})
    # print(date)
    for value2 in value.values():
        result[-1]['Чел'] += 1
        if value2[0]['type'] == 40091:
            result[-1]['Кратность человек'] += 1
        for value3 in value2:
            result[-1][value3['type']] += value3['count']
df = pd.DataFrame(result)
df.to_excel('result.xlsx', index=False)
