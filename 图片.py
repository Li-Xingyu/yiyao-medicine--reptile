# -*- coding: UTF-8 -*-
import requests
import random
from lxml import etree
from openpyxl import load_workbook
headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/55.0.2883.87 Safari/537.36'
            }
wb = load_workbook("1.xlsx")
sheet = wb["Sheet1"]

###################################


# 代理服务器
proxyHost = "n5.t.16yun.cn"
proxyPort = "6441"

# 代理隧道验证信息
proxyUser = "16UQASHB"
proxyPass = "259670"

proxyMeta = "http://%(user)s:%(pass)s@%(host)s:%(port)s" % {
    "host": proxyHost,
    "port": proxyPort,
    "user": proxyUser,
    "pass": proxyPass,
}

# 设置 http和https访问都是用HTTP代理
proxies = {
    "http": proxyMeta,
    "https": proxyMeta,
}

#  设置IP切换头
tunnel = random.randint(1, 10000)
headers = {"Proxy-Tunnel": str(tunnel),'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/55.0.2883.87 Safari/537.36'}
print proxies
###################################
for i in range(2,731):#2754
    try:
        html=etree.HTML(requests.get('http://http://yao.xywy.com/search/?q='+ sheet["B"+str(i)].value,headers=headers,timeout=5).text)
        img=(html.xpath('//table/tbody/tr[1]/td[4]/a/div/text()'))
    except:
        try:
            html = etree.HTML(requests.get('http://http://yao.xywy.com/search/?q=' + sheet["B" + str(i)].value, headers=headers, timeout=5).text)
            img=(html.xpath('//table/tbody/tr[1]/td[4]/a/div/text()'))
        except:
            pass
    if len(img)==0:
        try:
            html = etree.HTML(
                requests.get('http://http://yao.xywy.com/search/?q=' + sheet["B" + str(i)].value, headers=headers, timeout=5).text)
            img = (html.xpath('//table/tbody/tr[2]/td[4]/a/div/text()'))
        except:
            try:
                html = etree.HTML(
                    requests.get('http://http://yao.xywy.com/search/?q=' + sheet["B" + str(i)].value, headers=headers, timeout=5).text)
                img = (html.xpath('//table/tbody/tr[2]/td[4]/a/div/text()'))
            except:
                pass
    if len(img)==0:
        try:
            html = etree.HTML(
                requests.get('http://http://yao.xywy.com/search/?q=' + sheet["B" + str(i)].value, headers=headers, timeout=5).text)
            img = (html.xpath('//table/tbody/tr[3]/td[4]/a/div/text()'))
        except:
            try:
                html = etree.HTML(
                    requests.get('http://http://yao.xywy.com/search/?q=' + sheet["B" + str(i)].value, headers=headers, timeout=5).text)
                img = (html.xpath('//table/tbody/tr[3]/td[4]/a/div/text()'))
            except:
                pass
    print i
    x = 1
    imglist = []
    imglist=str(''.join(img)).split(',')
    if imglist[0]!="" :
        for k in range(0,len(imglist)):
            try:
                path = 'D:/pythone study/goods/'+ str(sheet["A"+str(i)].value) +'/%s.jpg' % x
                print imglist[k]
                try:
                    pic = requests.get(imglist[k], headers=headers,proxies=proxies,timeout=5)
                except:
                    try:
                        pic = requests.get(imglist[k], headers=headers, proxies=proxies,timeout=5)
                    except:
                        pass
                with open(path, 'wb') as f:
                    f.write(pic.content)
                    print("文件保存成功...")
                    x = x + 1
            except:
                print("文件保存失败...")
        img=""
        imglist = []
    else:
        print ("无照片...")
        with open('D:/pythone study/goods/1.txt', "a") as f:
            f.write("无照片......."+str(i))
            f.write('\n')

