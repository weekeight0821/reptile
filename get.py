from urllib import request,parse
import  xlrd
import xlwt
import re
from openpyexcel import Workbook
from bs4 import BeautifulSoup

def getHtml(page=0):
    url = "http://www.mafengwo.cn/gonglve/"
    header = {
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_3) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/74.0.3729.131 Safari/537.36',
        'X-Requested-With': 'XMLHttpRequest'
    }
    data = bytes(parse.urlencode({'page': page}), encoding='utf8')
    req = request.Request(url, data=data, headers=header)
    html = request.urlopen(req)
    code = html.getcode()
    if isinstance(code, int) and code == 200:
        return getdata(html)
    else:
        return 'error'


def getdata(html):

    soup = BeautifulSoup(html)

    row = []
    items = soup.find_all('div', class_='feed-item _j_feed_item')
    for item in items:
        content = item.contents[1]
        title = content.find('div', class_='title').get_text().strip()
        url = content.get('href')
        # stat = int(content.find('span', class_='num').get_text())
        info = content.find('div', class_='info').get_text().strip()
        source = content.find('strong').string
        author = ''
        comment = 0
        buystat = 0
        travelstat = 0
        repcomment = 0
        responsestat = 0
        ext_r = 0
        if source == '自由行攻略':
            buystat = int(content.find('span', class_='num').get_text())
            ext_r = int(content.find('li', class_='ext-r').get_text().replace('浏览', ''))
        elif source == '游记':
            travelstat = int(content.find('span', class_='num').get_text())
            ext_rs = content.find('span', class_='nums').get_text().split('，')
            ext_r = int(ext_rs[0].replace('浏览', ''))
            comment = int(ext_rs[1].replace('评论', ''))
            author = content.find('span', class_='author').get_text()
        elif source == '问答':
            responsestat = int(content.find('span', class_='num').get_text())
            ext_rs = content.find('span', class_='nums').get_text()
            ext_r = int(ext_rs[0].replace('浏览', ''))
            repcomment = int(ext_rs[1].replace('回答', ''))

        row.append([source, title, author, url, buystat,travelstat, responsestat,ext_r, comment,repcomment,info])
    return row

def wirtedata(data):
    wb = Workbook()
    dest_filename = 'test.xlsx'
    wsl = wb.active
    titlelist = ['来源', '标题', '作者', '文章的URL', '产品购买量', '游记点赞量', '问答点赞量', '浏览量', '评论数', '回答数', '文章简介']
    for row in range(len(titlelist)):
        c = row + 1
        wsl.cell(1, column=c, value=titlelist[row])
    for listindex in range(len(data)):
        for l in range(len(data[listindex])):
            wsl.append(data[listindex][l])
    wb.save(filename=dest_filename)
if __name__ == '__main__':
    rows = []
    for p in range(90):
        data = getHtml(p)
        rows.append(data)
    wirtedata(rows)
    print('ok')