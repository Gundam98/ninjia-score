# 操作excel的库
from openpyxl import load_workbook
import os
import sys
import json
import requests

import globalValue as glob

def refreshAccessToken():
    secretKey = 'vtk5zbBVi4hoVWclNwedp1orfxjtfkOM'
    apiKey = 'YmQ7Qi2Kooh5qOoENIAP1V4O'
    host = 'https://aip.baidubce.com/oauth/2.0/token?grant_type=client_credentials&client_id=' + apiKey + '&client_secret=' + secretKey
    response = json.loads(requests.get(host).text)
    print('get the access token: ' + response['access_token'])
    glob.access_token = response['access_token']

def getNameList(sheetName):
    inputPath = input('请输入表格名 (默认为: 胧月成绩统计总表.xlsx):')
    if inputPath != '': glob.excelPath = inputPath
    glob.excelPath = os.path.dirname(os.path.realpath(sys.argv[0])) + '/' + glob.excelPath
    excel = load_workbook(glob.excelPath)
    FuBen = excel[sheetName]
    nameCells = FuBen['B']
    nameList = []
    for cell in nameCells :
        if (cell.value != 'id') and (cell.value != None) and cell.fill.start_color.index != 1 :
            nameList.append(str(cell.value))
    return nameList