#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Author  : nacker 648959@qq.com
# @Time    : 2020/3/4 6:25 下午
# @Site    : 
# @File    : douban.py
# @Software: PyCharm

# 1.将目标网站上的页面抓取下来
# 2.将抓取下来的数据根据一定的规则进行提取
import time
import requests
import xlwt
from lxml import etree
from MysqlHelper import MysqlHelper

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/72.0.3626.119 Safari/537.36',
    'Referer': 'https://www.douban.com/',
}

def get_page_source(url):
    '''
    获取网页源代码
    # response = requests.get(url, headers=headers)
    # text = response.text
    # html = etree.HTML(text)
    '''
    try:
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            text = response.text
            return etree.HTML(text)
    except requests.ConnectionError:
        return None

def get_movie_info(html):
    if len(html):
        movies = []
        # item节点
        items = html.xpath("//div[@class='item']")
        for index,item in items:
            # 1.标题
            title = item.xpath("..//@alt")[0]
            # print(title)
            # 2.分数 rating_num
            score = item.xpath('..//span[@class="rating_num"]/text()')[0]
            # print(score)
            # 3.描述
            duration = ""
            # 4.时间+地区+类别
            tempString = item.xpath('div[@class="bd"]/p[@class=""]/text()')[1]
            tempString = "".join(tempString.split())
            temps = tempString.split('/')
            # print(temps)
            date = temps[0]
            # print(year)
            region = temps[1]
            # print(region)
            category = temps[2]
            # 5.导演+演员
            directorAndActorStr = item.xpath('div[@class="bd"]/p[@class=""]/text()')[0]
            directorAndActorStr = "".join(directorAndActorStr.split())
            # print(directorAndActorStr)
            directorAndActor = directorAndActorStr
            # 6.图片
            thumbnail = item.xpath("..//@src")[0]
            # print(thumbnail)
            # 7.quote
            quote = item.xpath('..//p[@class="quote"]/span/text()')
            if quote:
                quote = quote[0]
            else:
                quote = ""
            # print(quote)
            movie = {
                'title': title,
                'score': score,
                'date' : date,
                'region': region,
                'category': category,
                'directorAndActor': directorAndActor,
                'quote' : quote,
                'thumbnail': thumbnail
            }
            movies.append(movie)
        return movies


def writemovie(list):
    for dict in list:
        sql = 'insert into tb_movie(title,score,date,region,category,directorAndActor,quote,thumbnail) values(%s,%s,%s,%s,%s,%s,%s,%s)'
        mysqlHelper = MysqlHelper('localhost', 3306, 'douban', 'root', '123456')
        params = [dict['title'],dict['score'],dict['date'],dict['region'],dict['category'],dict['directorAndActor'],dict['quote'],dict['thumbnail']]
        # print(params)
        count = mysqlHelper.insert(sql, params)
        if count == 1:
            print(dict)
            print('--------------------ok--------------------')
        else:
            print('--------------------error--------------------')

# 将相关数据写入excel中
def saveToExcel(datalist):
    # 初始化Excel
    w = xlwt.Workbook()
    style = xlwt.XFStyle()  # 初始化样式
    font = xlwt.Font()  # 为样式创建字体
    font.name = u"微软雅黑"
    style.font = font  # 为样式设置字体
    ws = w.add_sheet(u"豆瓣电影Top250", cell_overwrite_ok=True)

    # 将 title 作为 Excel 的列名
    title = u"排行, 电影, 评分, 年份, 地区, 类别, 导演主演, 评价, 图片"
    title = title.split(",")
    for i in range(len(title)):
        ws.write(0, i, title[i], style)
    #  # 开始写入数据库查询到的数据
    for i in range(len(datalist)):
        row = datalist[i]
        for j in range(len(row)):
            if row[j]:
                item = row[j]
                ws.write(i + 1, j, item, style)

    # 写文件完成，开始保存xls文件
    path = '豆瓣电影Top250.xls'
    w.save(path)

def saveToMysql():
    # https: // movie.douban.com / top250?start = 0 & filter =
    for offset in range(0, 250, 25):
        url = 'https://movie.douban.com/top250?start=' + str(offset) +'&filter='
        item = get_page_source(url)
        # 获取单页数据
        list = get_movie_info(item)
        # 写数据到数据库
        writemovie(list)
        time.sleep(3)
    else:
        print('豆瓣top250的电影信息写入完毕')

def readMysqlData():
    sql = 'select * from tb_movie order by id asc'
    mysqlHelper = MysqlHelper('localhost', 3306, 'douban', 'root', '123456')
    datalist = mysqlHelper.get_all(sql)
    return datalist

def main():
    # 保存到数据库
    # saveToMysql()
    # 读取数据
    datalist = readMysqlData()
    # 写入Excel
    saveToExcel(datalist)

if __name__ == '__main__':
    main()
