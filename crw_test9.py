# -*- coding: utf-8 -*-
import xlwt
import urllib2
from tqdm import tqdm
from bs4 import BeautifulSoup
from operator import eq
import re


def bidnum_reader(name):
    a = []
    f = open(name + ".txt", 'rb').readlines()
    for item in f:
        b = item.decode('utf-8')
        a.append(b[:14])
    return a

bidnum_list = bidnum_reader('bidnum')
strlist = []

for i in bidnum_list:
    seq = '-'
    bidno = str(i).split(seq, 1)[0]
    strlist.append(bidno)

workbook = xlwt.Workbook(encoding='utf-8')
workbook.default_style.font.heignt = 20 * 11

xlwt.add_palette_colour("lightgray", 0x21)
workbook.set_colour_RGB(0x21, 216, 216, 216)
xlwt.add_palette_colour("lightgreen", 0x22)
workbook.set_colour_RGB(0x22, 216, 228, 188)

worksheet = workbook.add_sheet('sheet1')
col_width_0 = 256 * 13
col_width_1 = 256 * 13

col_height_content = 48

worksheet.col(0).width = col_width_0
worksheet.col(1).width = col_width_1

list_style = "font:height 180,bold on; pattern: pattern solid, fore_color lightgray; align: wrap on, vert centre, horiz center"

worksheet.write(0, 0, "계약명", xlwt.easyxf("font:height 180, bold on;pattern: pattern solid, fore_color lightgreen; align:vert centre, horiz center"))

row_marker = 0
column_marker = 0



for i in range(54,302):
    urlString = "http://gyeyak.chungnam.net/gyeyak/cntrt/cstnNgCntList.do?menuFg=C5&cntrtFlag=2&scode=1&pageIndex=" + str(i)
    body = urllib2.urlopen(urlString)
    soup = BeautifulSoup(body)
    parse_tr = soup.find_all('tbody')[1]

    columns = []

    for row in parse_tr.find_all('tr'):
        print("Page Number:" + str(i))
        columns = row('td', {'class': 'left'})[1]


        for column in columns:
            result1 = columns.get_text()
            column = re.compile(r'\s+')
            result2 = re.sub(column, ' ', result1)
            result3 = re.sub('[-=+,#/\?:^$.@*\"※~&%ㆍ!』\\‘|\(\)\\\<\>`\'…》]','', result2)
            result4 = re.sub(' ', '', result3)

            print(result4)
            worksheet.write(row_marker + 1, column_marker + 0, unicode(result4))
            break

        if len(columns) > 0:
            row_marker += 1


workbook.save('contractName_choongNam.xls')

print('##################### Finish ##################### ')


