#!/usr/bin/env python
# -*- coding: utf-8 -*-

import json
import xlwt
import requests
from datetime import datetime

url = 'http://yery.bjchyedu.cn/apix//nurserymanage/queryNurseryInfos'
body = {"communityCode": "", "streetCode": "QII5vwCWkg0rmsxBTKL4wDiCRkEHEPFF2sJwwJpAvczPBSTHD9kaqNrw+V7scxz6P6ZyABiaoGPAw7FiVghvJIQBZlYYl1aPhZTeGzud+KsvfWpag2zW97S1lJiPPXc5zA4co+M8gMGqgHfoU0U5Ixf8JEdWfiCBMqw3H+YN3/o="}
headers = {'content-type': "application/x-www-form-urlencoded; charset=UTF-8", 'Accept': "*/*", 'Accept-Encoding': "gzip, deflate", 'Accept-Language': "zh-CN,zh;q=0.9", 'Connection': "keep-alive", 'Content-Length': "210", 'Cookie': "BIGipServerchaoyangyouerruyuandengji80=788530092.20480.0000", 'Host': "yery.bjchyedu.cn", 'Origin': "http://yery.bjchyedu.cn/html/map.html", 'User-Agent': "Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/74.0.3729.131 Mobile Safari/537.36", 'X-Requested-With': "XMLHttpRequest"}

response = requests.post(url, data=body, headers=headers)
result = response.text
line = 1

style0 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on', num_format_str='#,##0.00')
style1 = xlwt.easyxf(num_format_str='D-MMM-YY')
wb = xlwt.Workbook(encoding = 'utf-8')
ws = wb.add_sheet('ChaoYang')

ws.write(0, 0, '幼儿园名称', style0)  # nurseryName
ws.write(0, 1, '幼儿园性质', style0)  # inclusiveClassFlagLabel
ws.write(0, 2, '幼儿园价格', style0)  # monthCostBa
ws.write(0, 3, '幼儿园等级', style0)  # nurseryLevelLabel
ws.write(0, 4, '幼儿园隶属', style0)  # nurseryNatureLabel
ws.write(0, 5, '幼儿园所属街道', style0)  # streetCodeName
ws.write(0, 6, '幼儿园联系电话', style0)  # enrollTelephone
ws.write(0, 7, '幼儿园地址', style0)  # detailedAddress
ws.write(0, 8, '幼儿园咨询时间', style0)  # askTime
ws.write(0, 9, '幼儿园简介', style0)  # nurseryCharacteristic

result = json.loads(result)
for street in result['result'].values():
    for nursery in street:
        nurseryName = nursery.get('nurseryName', '') if nursery.get('nurseryName', '') is not None else ''
        inclusiveClassFlagLabel = nursery.get('inclusiveClassFlagLabel', '') if nursery.get('inclusiveClassFlagLabel', '') is not None else ''
        monthCostBa = nursery.get('monthCostBa', '') if nursery.get('monthCostBa', '') is not None else ''
        nurseryLevelLabel = nursery.get('nurseryLevelLabel', '') if nursery.get('nurseryLevelLabel', '') is not None else ''
        nurseryNatureLabel = nursery.get('nurseryNatureLabel', '') if nursery.get('nurseryNatureLabel', '') is not None else ''
        streetCodeName = nursery.get('streetCodeName', '') if nursery.get('streetCodeName', '') is not None else ''
        enrollTelephone = nursery.get('enrollTelephone', '') if nursery.get('enrollTelephone', '') is not None else ''
        detailedAddress = nursery.get('detailedAddress', '') if nursery.get('detailedAddress', '') is not None else ''
        askTime = nursery.get('askTime', '') if nursery.get('askTime', '') is not None else ''
        nurseryCharacteristic = nursery.get('nurseryCharacteristic', '') if nursery.get('nurseryCharacteristic', '') is not None else ''
        ws.write(line, 0, nursery.get('nurseryName', ''))
        ws.write(line, 1, nursery.get('inclusiveClassFlagLabel', ''))
        ws.write(line, 2, nursery.get('monthCostBa', ''))
        ws.write(line, 3, nursery.get('nurseryLevelLabel', ''))
        ws.write(line, 4, nursery.get('nurseryNatureLabel', ''))
        ws.write(line, 5, nursery.get('streetCodeName', ''))
        ws.write(line, 6, nursery.get('enrollTelephone', ''))
        ws.write(line, 7, nursery.get('detailedAddress', ''))
        ws.write(line, 8, nursery.get('askTime', ''))
        ws.write(line, 9, nursery.get('nurseryCharacteristic', ''))
        line += 1

wb.save('nursery.xls')
