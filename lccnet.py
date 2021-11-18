# -*- coding: utf-8 -*-
"""
Created on Wed Sep 15 21:41:07 2021

@author: USER
"""

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook

wb = Workbook()
ws = wb.active

session_requests = requests.session()

url = 'https://member.lccnet.com.tw/login.asp?ACT=login'
header = {'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/93.0.4577.82 Safari/537.36'}
param = {
        'NO':'105363371',
        'PWD':'auo780502'
         }
content = session_requests.post(url, data = param, headers = header)
#print(content.text)

data = session_requests.get('https://member.lccnet.com.tw/myclass_index.asp', headers = header, allow_redirects = True)
data.encoding = 'cp950'
data = data.text
sp = BeautifulSoup(data, 'html.parser')

#已登記課程#
lessions = sp.find(id='table84')
trs = lessions.find_all('tr')

ws.append(['▼已登記課程'])

for row in trs:
    trsList = []
    tds = row.find_all('td')
    trsList.append(tds[0].text)
    trsList.append(tds[1].text.strip().replace(" ", ''))
    trsList.append(tds[2].text)
    trsList.append(tds[3].text)
    trsList.append(tds[4].text)
    ws.append(trsList)
    
#上課記錄#
lessions = sp.find(id='table85')
trs = lessions.find_all('tr')

ws.append(['▼上課記錄'])

for row in trs:
    trsList = []
    tds = row.find_all('td')
    trsList.append(tds[0].text)
    trsList.append(tds[1].text.strip().replace(" ", ''))
    trsList.append(tds[2].text)
    trsList.append(tds[3].text)
    trsList.append(tds[4].text.strip())
    ws.append(trsList)

wb.save('lession.xlsx')




