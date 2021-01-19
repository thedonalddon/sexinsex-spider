# -*- coding = utf-8 -*-
# @Time : 2021/1/17 14:21
# @Author : Donald
# @File : spider.py
# @Software : PyCharm

from bs4 import BeautifulSoup
import re
import urllib.request
import urllib.error
import xlwt
# import sqlite3

# 定义全局变量和搜索方法
threadlist = []
findLink = re.compile(r'<a href="(.*?)">')   # 正则表达式搜索链接
findImgSrc = re.compile(r'<img alt="" border="0" onclick="zoom(.*?)<br/>')
findImgSrc1 = re.compile(r'src="(.*?).jpg"/>')
findJpg = re.compile(r'src="(.*?).jpg"')
findJpeg = re.compile(r'src="(.*?).jpeg"')
# findTitle = re.compile(r'<h2>(.*?)"</h2>')
findTitle = re.compile(r'</a>(.*?)</h1>')
findTorrent = re.compile(r'<a href="attachment.php\?aid=(.*?)" target=')

baseurl = input("输入要爬取第地址（如http://sexinsex.net/bbs/forum-143-）：") # 填入爬取地址
# pages = 15
pages = input('从第2页开始爬，要爬到几页？')  #填入页数
pages = int(pages)


def main():
    threadlist = getThreadList(baseurl)
    content = getContent()
    savepath = "result.xls"
    saveData(content, savepath)


# 爬帖子列表
def getThreadList(baseurl):
    print("开始爬取...")
    for i in range(2, pages+1):
        url = baseurl + str(i) + '.html'
        print('开始爬取第%d页帖子列表，地址为' % i, url)
        html = askURL(url)

        # 逐一解析数据
        soup = BeautifulSoup(html, "html.parser")
        for item in soup.select("tr > th.lock > span > a"):
            item = str(item)
            link = re.findall(findLink, item)[0]
            threadlist.append('http://sexinsex.net/bbs/' + link)
    # print(threadlist)
    print('共找到%d个内容页' % len(threadlist))
    return threadlist


# 爬详情页
def getContent():
    content = []
    i = 0
    while i <= len(threadlist)-1:
        print('开始爬取第%d个内容页' % i)
        html = askURL(threadlist[i])
        # 逐一解析数据
        soup = BeautifulSoup(html, "html.parser")
        try:
            for item in soup.select("div > form"):
                item = str(item)
                # print(item)
                title = re.findall(findTitle, item)[0]          # 获取标题
                content.append(title)
                torrent = re.findall(findTorrent, item)[0]      # 获取种子
                content.append('http://sexinsex.net/bbs/attachment.php?aid='+torrent)
        except Exception as result:
            print('Exception occurs:')
            print(result)
        try:
            for item1 in soup:
                item1 = str(item1)
                # print(item1)
                jpg = re.findall(findJpg, item1)                  # 获取jpg图片
                # print(jpg)
                if jpg:                                           # 除空
                    content.append(jpg[0]+'.jpg')
                else:
                    continue
        except Exception as result:
            print('Exception occurs:')
            print(result)
        try:
            for item1 in soup:
                item1 = str(item1)
                # print(item1)
                jpeg = re.findall(findJpeg, item1)                # 获取jpg图片
                # print(jpeg)
                if jpeg:                                          # 除空
                    content.append(jpeg[0]+'.jpeg')
                else:
                    continue
        except Exception as result:
            print('Exception occurs:')
            print(result)
        i += 1
    # print(content)
    return content


# 伪装并爬取URL内容
def askURL(url):
    head = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.141 Safari/537.36"
        }
    request = urllib.request.Request(url, headers=head)
    html = ""
    try:
        response = urllib.request.urlopen(request)
        html = response.read().decode('gbk')
        # print(html)

    except urllib.error.URLError as e:
        if hasattr(e, "code"):
            print(e.code)
        if hasattr(e, "reason"):
            print(e.reason)
    return html


# 保存数据
def saveData(content, savepath):
    book = xlwt.Workbook(encoding="utf-8")
    sheet = book.add_sheet('帖子列表', cell_overwrite_ok=True)
    for i in range(0, len(content)):
        print('第%d条数据写入中' % i)
        sheet.write(i, 0, content[i])
    book.save('result.xls')
    # 批量下载图片


if __name__ == "__main__":
    main()
