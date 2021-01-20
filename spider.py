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
import datetime
import xlrd
import xlsxwriter

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

baseurl = input("输入要爬取的地址（如亚洲原创http://sexinsex.net/bbs/forum-143-）（或亚洲转帖http://sexinsex.net/bbs/forum-25-）：")  # 填入爬取地址
# pages = 15
pages = input('从第2页开始爬，要爬到几页？')  # 填入页数
pages = int(pages)

xlsname = ''
now = str(datetime.datetime.now())
xls = now[:-7] + '.xls'
def main():
    threadlist = getThreadList(baseurl)
    content = getContent()
    savepath = xls
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
        jindu = i *100 // len(threadlist)
        print('开始爬取第', i, '个内容页，', '共', len(threadlist), '个内容页，进度：', jindu, '%')
        html = askURL(threadlist[i])
        content.append(threadlist[i])
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
            print('title/torrent Exception occurs:')
            print(result)
        try:
            jpg = re.findall(findJpg, item)                         # 获取jpg图片
            # print(jpg)
            if jpg:                                                 # 除空
                q = 0
                while q < len(jpg):
                    content.append(jpg[q] + '.jpg')
                    q += 1
        except Exception as result:
            print('jpg Exception occurs:')
            print(result)
        try:
            jpeg = re.findall(findJpeg, item)                         # 获取jpeg图片
            # print(jpg)
            if jpeg:                                                 # 除空
                q = 0
                while q < len(jpeg):
                    content.append(jpeg[q] + '.jpg')
                    q += 1
        except Exception as result:
            print('jpeg Exception occurs:')
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
    book.save(savepath)

    booka = xlrd.open_workbook(savepath, 'r')  # 打开.xlsx文件
    sheeta = booka.sheets()[0]  # 打开表格中第一个sheet
    editedxls = 'edited-'+savepath
    bookb = xlsxwriter.Workbook(editedxls)  # 打开新xlsx文件
    sheetb = bookb.add_worksheet('帖子列表')

    nrows = sheeta.nrows  # 获取总行数
    rowAlist = []

    for i in range(0, nrows):
        rowAcontent = sheeta.cell(i, 0).value  # 遍历A列所有数据
        sheetb.write(i, 0, rowAcontent)
        if 'jpg' in rowAcontent:
            sheetb.write(i, 1, rowAcontent)
    bookb.close()

    print('数据已保存为 %s 。Congrats! 下一步：执行convertimg.py ' % editedxls)
    # 批量下载图片









if __name__ == "__main__":
    main()
