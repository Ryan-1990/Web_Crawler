# -*- coding: utf-8 -*-
#!/usr/bin/env python  
#抓取百度搜索结果  
import sys
import re
import time
import urllib2
import xlrd
from xlutils.copy import copy

def search(key):
    search_url='http://www.baidu.com/s?wd=key&rsv_bp=0&rsv_spt=3&rsv_n=2&inputT=6391'
    req=urllib2.urlopen(search_url.replace('key',key))
    html=req.read()
    href_temp=re.findall('http://opendata.baidu.com/yaopin/\S*utf-8"\s*[\w=_>"]*\s*查看全部相关药品', html)
    #print href_temp
    href=re.search('http://opendata.baidu.com/yaopin/\S*utf-8', str(href_temp))
    return href


fname = "1.xls"
bk = xlrd.open_workbook(fname, formatting_info=True)
shxrange = range(bk.nsheets)
try:
    sh = bk.sheet_by_name("Sheet1")
except:
    print "no sheet in %s named Sheet1" % fname

nrows = sh.nrows #获取行数
wb = copy(bk)
ws = wb.get_sheet(0)

#for i in range(1,10):
for i in range(1000,nrows):
    #start_time = time.time()
    key=sh.cell_value(i,0).encode('utf8')
    print key
    result=search(key)

    if result:
        print result.group()
        ws.write(i, 8, result.group())
    else:
        print "Error"
        ws.write(i, 8, "Error")
    #print "Time elapsed: ", time.time() - start_time, "s"
    if i%100==0:
        wb.save(fname)
wb.save(fname)
