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
import os
import requests
from PIL import Image
import shutil

# 定义全局变量和搜索方法
threadlist = []
xlsname = ''
now = str(datetime.datetime.now())
xls = now[:-7] + '.xls'
savepath = xls
findLink = re.compile(r'<a href="(.*?)">')   # 正则表达式搜索链接
findImgSrc = re.compile(r'<img alt="" border="0" onclick="zoom(.*?)<br/>')
findImgSrc1 = re.compile(r'src="(.*?).jpg"/>')
findJpg = re.compile(r'src="(.*?).jpg"')
findJpeg = re.compile(r'src="(.*?).jpeg"')
findTitle = re.compile(r'</a>(.*?)</h1>')
findTorrent = re.compile(r'<a href="attachment.php\?aid=(.*?)" target=')

baseurl = input("输入要爬取的地址（如亚洲原创http://sexinsex.net/bbs/forum-143-）（或亚洲转帖http://sexinsex.net/bbs/forum-25-）：")  # 填入爬取地址
startpage = input('从第几页开始爬？Starting page? （建议至少为2）') # 填入起始页
startpage = int(startpage)
pages = input('从第 %d 页开始爬，要爬到几页？Ending page? ' % startpage)  # 填入页数
pages = int(pages)


def main():
    threadlist = getThreadList(baseurl)     # 爬帖子列表，提取内容页链接
    content = getContent()                  # 爬内容页，提取有效数据
    saveData(content, savepath)             # 保存有效数据到xls
    downloadImg()                           # 下载图片并加入xls
    os.remove(xls)                          # 清空缓存


# 爬帖子列表
def getThreadList(baseurl):
    print("开始爬取...")
    for i in range(startpage, pages+1):
        url = baseurl + str(i) + '.html'
        print('开始爬取第%d页帖子列表，地址为' % i, url)
        html = askURL(url)

        # 解析数据并提取内容页链接
        soup = BeautifulSoup(html, "html.parser")
        for item in soup.select("tr > th.lock > span > a"):
            item = str(item)
            link = re.findall(findLink, item)[0]
            threadlist.append('http://sexinsex.net/bbs/' + link)
    print('共找到%d个内容页' % len(threadlist))
    return threadlist


# 爬详情页
def getContent():
    content = []
    i = 0
    while i <= len(threadlist)-1:
        jindu = i * 100 // len(threadlist)
        print('开始爬取第', i, '个内容页，', '共', len(threadlist), '个内容页，进度：', jindu, '%')
        html = askURL(threadlist[i])
        content.append(threadlist[i])
        # 解析数据，提取标题、种子、图片
        soup = BeautifulSoup(html, "html.parser")
        try:
            for item in soup.select("div > form"):
                item = str(item)
                title = re.findall(findTitle, item)[0]          # 获取标题
                content.append(title)
                torrent = re.findall(findTorrent, item)[0]      # 获取种子
                content.append('http://sexinsex.net/bbs/attachment.php?aid='+torrent)
        except Exception as result:
            print('title/torrent Exception occurs:')
            print(result)
        try:
            jpg = re.findall(findJpg, item)                         # 获取jpg图片
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
            if jpeg:                                                 # 除空
                q = 0
                while q < len(jpeg):
                    content.append(jpeg[q] + '.jpg')
                    q += 1
        except Exception as result:
            print('jpeg Exception occurs:')
            print(result)
        i += 1
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

    except urllib.error.URLError as e:
        if hasattr(e, "code"):
            print(e.code)
        if hasattr(e, "reason"):
            print(e.reason)
    return html


# 保存有效数据到xls
def saveData(content, savepath):
    book = xlwt.Workbook(encoding="utf-8")
    sheet = book.add_sheet('帖子列表', cell_overwrite_ok=True)
    for i in range(0, len(content)):
        sheet.write(i, 0, content[i])
    book.save(savepath)

    # 复制A列jpg到B列
    booka = xlrd.open_workbook(savepath, 'r')  # 打开.xlsx文件
    sheeta = booka.sheets()[0]  # 打开表格中第一个sheet
    editedxls = 'edited-'+savepath
    bookb = xlsxwriter.Workbook(editedxls)  # 打开新xlsx文件
    sheetb = bookb.add_worksheet('帖子列表')

    nrows = sheeta.nrows  # 获取总行数
    # rowAlist = []

    for i in range(0, nrows):
        rowAcontent = sheeta.cell(i, 0).value  # 遍历A列所有数据
        sheetb.write(i, 0, rowAcontent)
        if 'jpg' in rowAcontent:
            sheetb.write(i, 1, rowAcontent)
    bookb.close()
    print('爬取完成！数据已保存在： %s 。准备下载所有图片... ' % editedxls)


# 下载editedxls里的所有图片，并放进同名文件夹
def downloadImg():
    editedxls = 'edited-'+savepath
    filename = editedxls
    newdir = filename.strip('.xls')
    a = xlrd.open_workbook(filename, 'r')  # 打开.xlsx文件
    sht = a.sheets()[0]  # 打开表格中第一个sheet
    nrows = sht.nrows  # 获取总行数

    path = os.getcwd()
    isExists = os.path.exists(newdir)  # 新建同名文件夹
    if isExists:
        shutil.rmtree(newdir)
    os.mkdir(newdir)
    spath = os.path.join(path, newdir)
    spath1 = os.path.join(spath + '/')
    print('正在保存到目录：' + spath1)

    for i in range(0, nrows):
        url = sht.cell(i, 1).value  # 遍历B列所有数据
        if url != '':
            try:
                f = requests.get(url)
            except Exception as result:
                print('第 %d 行图片下载失败' % (i + 1))
                print(result)
                continue
            ii = str(i)  # 按照下载顺序（行号）构造文件名
            dir = ii + "." + 'jpg'  # 构造完整文件名称
            with open(spath1 + dir, "wb") as code:
                code.write(f.content)  # 保存文件
            print(url)  # 打印当前的 URL
            jindu = i * 100 // nrows  # 计算下载进度
            print('正在下载第 %d 张图片，共 %d 张图片，下载进度：%d' % (i + 1, nrows, jindu), '%')  # 显示下载进度
    print('所有图片下载完成！')
    print('*' * 30)

    # 将图片插入新的xls
    imglist = []
    imglist = os.listdir(spath1)
    booka = xlrd.open_workbook(savepath, 'r')  # 打开链接列表文件
    sheeta = booka.sheets()[0]  # 打开表格中第一个sheet

    book = xlsxwriter.Workbook('img_' + filename + 'x')  # 准备写入的新文件
    sheet = book.add_worksheet('img')
    sheet.set_column(1, 1, width=5000)

    nrows = sheeta.nrows  # 获取总行数
    # rowAlist = []

    for i in range(0, nrows):
        rowAcontent = sheeta.cell(i, 0).value  # 遍历booka.A列所有数据
        sheet.write(i, 0, rowAcontent)

    i = 0
    while i <= len(imglist) - 1:
        imgname = imglist[i]
        imgname = imgname[:-4]  # 去后缀
        imgnumber = str(imgname)
        place = 'B%s' % imgnumber
        imgfilename = os.path.join(spath1, imglist[i])
        i += 1
        try:
            with Image.open(imgfilename) as imgsize:
                imgheight = imgsize.height
            sheet.set_row((int(imgnumber) - 1), int(imgheight))
            sheet.insert_image(place, imgfilename)
            print('正在把 %s.jpg 添加进新的xls' % imgname)
        except Exception as result:
            print('添加失败')
            print(result)
            continue
    book.close()
    print('%d 张图片全部写入完成！结果保存在：%s ' % (nrows, imgfilename))
    os.remove(editedxls)


if __name__ == "__main__":
    main()
