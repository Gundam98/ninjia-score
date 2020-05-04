from PIL import Image
import requests
import os
import json
from urllib.parse import quote
import base64
import sys
import random

import globalValue as glob
from utils import refreshAccessToken



templateSignList = {
    1: {
        '16:9': '1f7f73f7e9b25b4918e02974a6683ce3',
        '18:9': 'be06072f8da2ccd50a7d8cd6b65c740b',
    },
    2: {
        '16:9': '95952d339513cf56bb46347ea14a561b',
        '18:9': '94c915b4d8589516742d92b936e9e717',
    },
    3: {
        '16:9': 'ff0122edea537f21e93e0a9334ad93dc',
        '18:9': 'd833009061f3d2260944a841b4cf9f32',
    },
    4: '',
}


'''
getImgInfo()的返回值的数据格式
副本:
    {
        name(string): score(string)
    }
深渊:
    {
        name(string): [damage(string), times(string), average(string)]
    }
33:
    {
        name(string): [win(boolean), rate(string)]
    }
'''

def getImgInfo(signId):
    print('图片格式必须为: pic+数字，请务必修改')
    photoType = input("请输入图片后缀 (默认为: png):")
    if photoType == '': photoType = 'png'

    result = {}
    i = 1
    while 1:
        imagepath = os.path.dirname(os.path.realpath(sys.argv[0])) + '/' + 'pic' + str(i) + '.' + photoType

        try:
            rawResult = getOCRResult(imagepath, templateSignList[signId])
            i += 1
        except Exception as e:
            if i == 1:
                print('未找到文件，请确认文件名是否正确')
                exit(1)
            else:
                break
        if signId == 1:
            result = {**result, **generatePreparationDungeonData(rawResult)}
        elif signId == 2:
            result = {**result, **generateAbyssData(rawResult)}
        elif signId == 3:
            result = {**result, **generateFamilyWarDungeonData(rawResult)}
        else:
            result = {**result, **generateFightData(rawResult)}

    print('数据如下：')
    i = 1
    for k, v in result.items():
        print(str(i)+ ': ' + k + '\t->\t', end='')
        print(v)
        i += 1
    return result

def getOCRResult(image_path, templateSign):
    try:
        recognize_api_url = "https://aip.baidubce.com/rest/2.0/solution/v1/iocr/recognise"
        image = get_file_content(image_path)
        print('识别中...', image_path)
        img = Image.open(image_path)
        (picW, picH) = img.size
        if picW / picH > 1.87 and picW / picH <= 2.2:
            templateSign = templateSign['18:9']
        elif picW / picH >= 1.67 and picW / picH <= 1.87:
            templateSign = templateSign['16:9']
        else:
            print('当前图片的尺寸无法识别')
            exit(3)
        recognize_body = "access_token=" + glob.access_token + "&templateSign=" + templateSign + "&image=" + quote(image)
        headers = {
            'Content-Type': "application/x-www-form-urlencoded",
            'charset': "utf-8"
        }
        response = requests.post(recognize_api_url, data=recognize_body, headers=headers)
        responseJson = json.loads(response.text)
        if responseJson['error_code'] == '111' or responseJson['error_code'] == '110':
            refreshAccessToken()
            recognize_body = "access_token=" + glob.access_token + "&templateSign=" + templateSign + "&image=" + quote(image)
            headers = {
                'Content-Type': "application/x-www-form-urlencoded",
                'charset': "utf-8"
            }
            response = requests.post(recognize_api_url, data=recognize_body, headers=headers)
            responseJson = json.loads(response.text)
        result = responseJson['data']
        if not result['isStructured']:
            print('OCR识别失败，请确认识别是否完整，具体规则可联系管理员获知')
            print('error message:',responseJson['errorMsg'])
            exit(2)
        return result['ret']

    except Exception as e:
        raise Exception('')

def get_file_content(file):
    try:
        with open(file, 'rb') as f:
            image_data = f.read()
        image_b64 = base64.b64encode(image_data)
        return image_b64
    except Exception as e:
        raise Exception('')

def generateAbyssData(rawData):
    j = 0
    result = {}
    while j < len(rawData):
        if 'member' in rawData[j]['word_name'] and 'damage' in rawData[j+1]['word_name'] and 'times' in rawData[j+2]['word_name']:
            name = rawData[j]['word']
            if "新人" in name or "族员" in name or "精英" in name or "豪杰" in name or "长老" in name or "副族长" in name or "族长" in name:
                name = name.replace('新人','')
                name =name.replace('族员','')
                name =name.replace('精英','')
                name =name.replace('豪杰','')
                name =name.replace('长老','')
                name =name.replace('副族长','')
                name =name.replace('族长','')
            damage = rawData[j+1]['word']
            times = rawData[j+2]['word']
            try:
                average = str(int(int(damage) / int(times)))
            except Exception as e:
                print('\033[0;37;41mERROR\033[0m 族员:%s的平均伤害计算失败,请留意是否识别出错。他的成绩为:%s,%s' % (name, damage, times))
                average = '0'
            if name == '':
                name = 'empty name#' + str(random.randint(1, 100) * random.randint(1, 100))
            result.update({name: [damage, times, average]})
        else:
            print('数据出错，请重试。')
        j += 3
    return result


def generatePreparationDungeonData(rawData):
    j = 0
    result = {}
    while j < len(rawData):
        if 'member' in rawData[j]['word_name'] and 'score' in rawData[j+1]['word_name']:
            name = rawData[j]['word']
            score = rawData[j+1]['word']
            if score == '':
                score = '0'
            if "新人" in name or "族员" in name or "精英" in name or "豪杰" in name or "长老" in name or "副族长" in name or "族长" in name:
                name = name.replace('新人','')
                name =name.replace('族员','')
                name =name.replace('精英','')
                name =name.replace('豪杰','')
                name =name.replace('长老','')
                name =name.replace('族长','')
                name =name.replace('副族长','')
            if name == '':
                name = 'empty name#' + str(random.randint(1, 100) * random.randint(1, 100))
            result.update({name: score})
        else:
            print('数据出错，请重试。')
        j += 2
    return result

def generateFamilyWarDungeonData(rawData):
    j = 0
    tempDict = {}
    result = {}
    # 手动提取json信息, 此时序号顺序可能不一致
    while j < len(rawData):
        tempDict.update({rawData[j]['word_name']:rawData[j]['word']})
        j += 1
    j = 1
    while j <= 6:
        name = 'member#' + str(j)
        score = 'score#' + str(j)

        name = tempDict[name]
        score = tempDict[score]
        if name == '':
            name = 'empty name#' + str(random.randint(1, 100) * random.randint(1, 100))
        result.update({name:score})
        j += 1
    return result

def generateFightData():
    return