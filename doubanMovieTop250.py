# -*- coding: utf-8 -*-

import requests
from bs4 import BeautifulSoup
import xlwt


# 请求
def request_douban(url):
    try:
        header = {
            'User-Agent': 'Mozilla / 5.0(Windows NT 10.0;Win64;x64) AppleWebKit / 537.36(KHTML, likeGecko) Chrome / 115.0.0.0Safari / 537.36'
        }
        response = requests.get(url, headers=header, timeout=10)
        if response.status_code == 200:
            return response.text
    except requests.RequestException as e:
        print('请求异常：' + e)
        return None


def run(page, sheet):
    url = 'https://movie.douban.com/top250?start=' + str(page * 25) + '&filter='
    html = request_douban(url)
    soup = BeautifulSoup(html, 'lxml')
    item_list = soup.find(class_='grid_view').find_all('li')

    for item in item_list:
        item_name = item.find(class_='title').string
        item_img = item.find('a').find('img').get('src')
        item_index = item.find(class_='').string
        item_score = item.find(class_='rating_num').string
        item_author = item.find('p').text.strip().split('\n')[0]
        item_intr = ''
        if item.find(class_='inq') is not None:
            item_intr = item.find(class_='inq').string
        print('爬取电影：' + item_index + ' | ' + item_name + ' | ' + item_score + ' | ' + item_intr)

        # 保存数据到指定行列位置
        sheet.write(int(item_index), 0, item_name)
        sheet.write(int(item_index), 1, item_img)
        sheet.write(int(item_index), 2, item_index)
        sheet.write(int(item_index), 3, item_score)
        sheet.write(int(item_index), 4, item_author)
        sheet.write(int(item_index), 5, item_intr)

    # 保存文件
    book.save(u'豆瓣最受欢迎的250部电影.xls')


if __name__ == '__main__':
    # 创建Excel文件，设置标题行
    book = xlwt.Workbook(encoding='utf-8', style_compression=0)
    sheet = book.add_sheet('豆瓣电影Top250', cell_overwrite_ok=True)
    sheet.write(0, 0, '名称')
    sheet.write(0, 1, '图片')
    sheet.write(0, 2, '排名')
    sheet.write(0, 3, '评分')
    sheet.write(0, 4, '作者')
    sheet.write(0, 5, '简介')

    for i in range(0, 10):
        run(i, sheet)
