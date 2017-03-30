# -*- coding: utf-8 -*-
#Author: Nan LI @ Tianjin University
#功能描述：遍历药源网的所有相关链接(西药和中药两类)，抓取网页内容存到本地

import sys
import time
import urllib2
import requests
import lxml.html
import random
import string
import re


# function: get the directory of script that calls this function
# usage: path = script_path()
def script_path():
    import inspect, os
    caller_file = inspect.stack()[1][1]         # caller's filename
    return os.path.abspath(os.path.dirname(caller_file))# path

def create_folder():
    import os
    path = script_path()
    title = "Yao_Yuan"
    new_path = os.path.join(path, title)
    if not os.path.isdir(new_path):
        os.makedirs(new_path)
        
    path += "/Yao_Yuan"
    title = "Xi_Yao"
    new_path = os.path.join(path, title)
    if not os.path.isdir(new_path):
        os.makedirs(new_path)
    title = "Zhong_Yao"
    new_path = os.path.join(path, title)
    if not os.path.isdir(new_path):
        os.makedirs(new_path)

class Search:
    def __init__(self, address):
        #print address
        self.address = address

        randnum = random.randint(1,9)
        if randnum == 1:
            self.header = {'Accept-Charset': 'utf-8', 'Accept': 'text/css,*/*;q=0.1',
                           'User-Agent': 'Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/33.0.1750.154 Safari/537.36'}
        elif randnum == 2:
            self.header = {'Accept-Charset': 'utf-8', 'Accept': 'text/css,*/*;q=0.1',
                           'User-Agent': 'Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.1; Trident/4.0)',
                           'Referer': 'http://www.baidu.com'}
        elif randnum == 3:
            self.header = {'Accept-Charset': 'utf-8', 'Accept': 'text/css,*/*;q=0.1',
                           'User-Agent': 'Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.1; Win64; x64; Trident/4.0)'}
        elif randnum == 4:
            self.header = {'Accept-Charset': 'utf-8', 'Accept': 'text/css,*/*;q=0.1',
                           'User-Agent': 'Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.1; Trident/4.0)'}
        elif randnum == 5:
            self.header = {'Accept-Charset': 'utf-8', 'Accept': 'text/css,*/*;q=0.1',
                           'User-Agent': 'Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.1; WOW64; Trident/4.0)'}
        elif randnum == 6:
            self.header = {'Accept-Charset': 'utf-8', 'Accept': 'text/css,*/*;q=0.1',
                           'User-Agent': 'Mozilla/5.0 (Windows; U; Windows NT 5.2) Gecko/2008070208 Firefox/3.0.1'}
        elif randnum == 7:
            self.header = {'Accept-Charset': 'utf-8', 'Accept': 'text/css,*/*;q=0.1',
                           'User-Agent': 'Mozilla/5.0 (Windows; U; Windows NT 5.1) Gecko/20070803 Firefox/1.5.0.12'}
        elif randnum == 8:
            self.header = {'Accept-Charset': 'utf-8', 'Accept': 'text/css,*/*;q=0.1',
                           'User-Agent': 'Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 5.1; .NET CLR 2.0.50727; Maxthon 2.0'}
        else:
            self.header = {'Accept-Charset': 'utf-8', 'Accept': 'text/css,*/*;q=0.1',
                           'User-Agent': 'Mozilla/5.0 (Windows; U; Windows NT 5.2) AppleWebKit/525.13 (KHTML, like Gecko) Version/3.1 Safari/525.13'}

        #sleepseconds = random.uniform(0, 2)   # Sleep randomly from 0 sec to 5.0 sec. before launching another crawling.
        #time.sleep(sleepseconds)

        #data = requests.get(self.address, headers=self.header).text   # This method is easy to be blocked.

        req = requests.Request(method='GET',
                               url=self.address,
                               headers=self.header,
                               cookies=None)
        reqprep = req.prepare()
        s = requests.Session()
        resp = s.send(reqprep)

        self.text = resp.text
        self.tree = lxml.html.fromstring(resp.text)
        
class ExtractorWebsite(Search):
    def extractor_website(self):
        a = self.tree.xpath('//title')
        if len(a) == 0:
            return False
        return a[0].text.encode('utf8')

def Start():
    create_folder()
    ScriptPath = script_path()
    XY_Path = ScriptPath+"/Yao_Yuan/Xi_Yao/"
    ZY_Path = ScriptPath+"/Yao_Yuan/Zhong_Yao/"

    X_address = "http://www.yaopinnet.com/huayao1/"
    Z_address = "http://www.yaopinnet.com/zhongyao1/"

    ############################西药######################################
    for word in string.lowercase:
        for index in xrange(1,30):
            X_address_URL = X_address+word+'%d'%index+'.htm'
            copy = ExtractorWebsite(X_address_URL)
            result = copy.extractor_website()
            if re.search('404',result):
                break
            spath = XY_Path+"XY_"+word+'%d'%index+".htm"
            f=open(spath,"w")
            f.write(copy.text.encode('utf8'))
            f.close()

    ############################中药######################################
    for word in string.lowercase:
        for index in xrange(1,30):
            Z_address_URL = Z_address+word+'%d'%index+'.htm'
            copy = ExtractorWebsite(Z_address_URL)
            result = copy.extractor_website()
            if re.search('404',result):
                break
            spath = ZY_Path+"ZY_"+word+'%d'%index+".htm"
            f=open(spath,"w")
            f.write(copy.text.encode('utf8'))
            f.close()


if __name__ == "__main__":
    Start()
