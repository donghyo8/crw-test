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


def get_bidurl(bidnum):
    num_split = str(bidnum).split(sep='-')
    bidno = num_split[0]
    return bidnum


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

worksheet.write(0, 0, "추정가격", xlwt.easyxf("font:height 180, bold on;pattern: pattern solid, fore_color lightgreen; align:vert centre, horiz center"))

row_marker = 0
column_marker = 0

for i in range(0, len(strlist)):
    urlString ="http://www.g2b.go.kr:8081/ep/invitation/publish/bidInfoDtl.do?bidno="+strlist[i]+"&bidseq=00&releaseYn=Y&taskClCd=1"
    body = urllib2.urlopen(urlString)
    soup = BeautifulSoup(body)

    print("공고번호:" + strlist[i])
    print(str(i) + "건")

    parse_tr = soup.find_all('table', {'class': 'table_info'})[3]

    tb_inner_count = 0

    for row in parse_tr.find_all('td'):
        column_marker = 0
        columns = row.find_all('div', {'class': 'tb_inner'})
        tb_inner_count+=1

        if tb_inner_count == 4:
            break

    for column in columns:
        result1 = column.get_text()
        column = re.compile(r'\s+')
        result2 = re.sub(column, ' ', result1)
        result3 = re.sub('[-=+,#/\?:^$.@*\"※~&%ㆍ!』\\‘|\(\)\[\]\<\>`\'…》]','', result2)
        result4 = re.sub(' ', '', result3)

        print(result4)
        worksheet.write(row_marker + 1, column_marker + 0, unicode(result4))
        column_marker += 1

    if len(columns) > 0:
        row_marker += 1


workbook.save('Test_20470101-20470331.xls')

print('##################### Finish ##################### ')


