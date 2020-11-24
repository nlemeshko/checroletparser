import requests
import json
import xlwt
from tempfile import TemporaryFile
import pandas as pd
import re
import translators as ts



column = list()
one = list()
one_rus = list()
two = list()
three = list()
four = list()
start=27
bolt = pd.read_csv('bolt-ev.csv', index_col=0)
pd.options.display.max_colwidth = 10000


def striphtml(data):
    p = re.compile(r'<.*?>')
    return p.sub('', data)

book = xlwt.Workbook()
sheet1 = book.add_sheet('Chevrolet')
sheet2 = book.add_sheet('Chevrolet_RUS')

i=0
for i in range(len(bolt.columns)):
    column.append(bolt.columns[i])

one.append('')
two.append(bolt[str(13)][1])
three.append(bolt[str(13)][2])
four.append('')

for i in range(3):
    one.append('')
    two.append(bolt[str(10+i)][1])
    three.append(bolt[str(10+i)][2])
    four.append('')



for x in range(len(bolt['18'])):
    try:
        one.append(striphtml(bolt['18'][x]))
    except Exception:
        one.append(bolt['18'][x])
    two.append(bolt[str(x+start)][1])
    three.append(bolt[str(x+start)][2])
    try:
        four.append(striphtml(bolt['23'][x]))
    except Exception:
        four.append(bolt['23'][x])

one.pop(4)
two.pop(4)
three.pop(4)
four.pop(4)

for i,e in enumerate(one):
    sheet1.write(i,0,e)
    try:
        sheet2.write(i,0,ts.alibaba(e, to_language='ru', sleep_seconds=0, use_cache=True))
        print(e)
    except Exception:
        sheet2.write(i, 0,e)
for i,e in enumerate(two):
    sheet1.write(i,1,e)
    sheet2.write(i, 1, e)
for i,e in enumerate(three):
    sheet1.write(i,2,e)
    sheet2.write(i, 2, e)
for i,e in enumerate(four):
    sheet1.write(i,3,e)
    try:
        sheet2.write(i, 3, ts.alibaba(e, to_language='ru', sleep_seconds=0, use_cache=True))
        print(e)
    except Exception:
        sheet2.write(i, 3, e)




name = "test.xls"
book.save(name)
book.save(TemporaryFile())



