# -*- coding: UTF-8 -*-
import requests
import random
import sys
reload(sys)
sys.setdefaultencoding('utf8')
from lxml import etree
from openpyxl import load_workbook
headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/55.0.2883.87 Safari/537.36'
            }
wb = load_workbook("1.xlsx")
sheet = wb["Sheet1"]
name=[]
wzd=0
for i in range(2,731):#2754
    try:
        html=etree.HTML(requests.get('https://www.315jiage.cn/search.aspx?where=title&keyword='+ sheet["B"+str(i)].value+ sheet["D"+str(i)].value+ sheet["E"+str(i)].value,headers=headers).content)
        html=etree.HTML(requests.get('https://www.315jiage.cn/'+html.xpath('//tr/td[1]/div[1]/div[2]/div/div[1]/div[1]/div[1]/a/@href')[0],headers=headers).content)
        a=  html.xpath('//*[@id="tab1"]/ul//text()')
        for j in a:
            with open('D:/pythone study/goods/'+ sheet["A"+str(i)].value +'/intro.txt', "a") as f:
                f.write(j.decode('UTF-8'))
                f.write('\n')

        print ('D:/pythone study/goods/' +  sheet["A"+str(i)].value + '/intro.txt')+'-----------------------'+ str(i)
    except:
        print '未找到'
        with open('D:/pythone study/goods/结果.txt', "a") as f:
            f.write("未找到......."+str(i))
            f.write('\n')
        wzd=wzd+1
print wzd