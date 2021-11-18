# -*- coding: utf-8 -*-
"""
Created on Wed Sep 15 19:53:41 2021

@author: USER
"""

import requests
import json
from openpyxl import Workbook
# from openpyxl import load_workbook

wb = Workbook()
ws = wb.active

ws.column_dimensions['B'].width = 18
ws.column_dimensions['C'].width = 18

url = 'https://www.thsrc.com.tw/TimeTable/Search'
header = {'user-agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/93.0.4577.82 Safari/537.36'}
param = {
        'SearchType': 'S',
        'Lang': 'TW',
        'StartStation': 'NanGang',
        'EndStation': 'ZuoYing',
        'OutWardSearchDate': '2021/09/18',
        'OutWardSearchTime': '18:00',
        'ReturnSearchDate': '2021/09/18',
        'ReturnSearchTime': '18:00'
        }

title = ["車次", "{}出發時間".format(param['StartStation']), "抵達{}時間".format(param['EndStation']), "旅行時間"]

data = requests.post(url, data = param, headers = header)
data.encoding = 'UTF-8'
data = data.text
thsrc = json.loads(data)

station = thsrc['data']['DepartureTable']['TrainItem']



for row in station:
    ws.append(title)
    number = []
    # startTime = []
    # endTime = []
    # duration = []
    number.append(row['TrainNumber'])
    number.append(row['DepartureTime'])
    number.append(row['DestinationTime'])
    number.append(row['Duration'])
    ws.append(number)
    ws.append(['各站停靠', '出發時間'])
    StationInfo = row['StationInfo']
    for s in StationInfo:
        if s['DepartureTime'] != "":
            stationName = []
            stationName.append(s['StationName'])
            stationName.append(s['DepartureTime'])
            ws.append(stationName)
    ws.append([''])

wb.save("thsrc.xlsx")




# highway = pd.DataFrame({
#                         "車次":number,
#                         "出發時間":startTime,
#                         "到達時間":endTime,
#                         "抵達時間":endTime,
#                         "旅行時間":duration
                        
# },columns=[
#             "車次",
#             "出發時間",
#             "到達時間",
#             "抵達時間",
#             "旅行時間"
            
# ])
    
    

# print(highway)

# print(stationName)






