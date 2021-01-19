# -*- coding: UTF-8 -*- 

import urllib.request
import xlwt
import requests
import os
import xlrd
import re
import shutil

a = xlrd.open_workbook('result-cg.xls', 'r')  # 打开.xlsx文件
sht = a.sheets()[0]  # 打开表格中第一个sheet
nrows = sht.nrows # 获取总行数

path = os.getcwd()
isExists = os.path.exists('convert')
if isExists:
    shutil.rmtree('convert')
os.mkdir('convert')

for i in range(0, nrows):
    url = sht.cell(i, 1).value # 遍历B行所有数据
    if url != '':
        f = requests.get(url)
        ii = str(i)  # 按照下载顺序（行号）构造文件名
        url2 = url[-3:]  # 根据链接地址获取文件后缀，后缀有.jpg 和 .gif 两种
        dir = ii + "." + url2  # 构造完整文件名称
        spath = os.path.join(path, 'convert/')
        with open(spath + dir, "wb") as code:
            code.write(f.content)  # 保存文件
        print(url)  # 打印当前的 URL
        jindu = i / nrows * 100  # 计算下载进度
        print("下载进度：", jindu, "%")  # 显示下载进度