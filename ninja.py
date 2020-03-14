from openpyxl import load_workbook
from openpyxl.styles import Color, PatternFill, Border,Side
from aip import AipOcr
from fuzzywuzzy import fuzz
import json
import os
import sys

config = {
    'appId': '18723556',
    'apiKey': 'YmQ7Qi2Kooh5qOoENIAP1V4O',
    'secretKey': 'vtk5zbBVi4hoVWclNwedp1orfxjtfkOM'
}

FBType = {
    '1': '蝙蝠',
    '2': '西瓜',
    '3': '金币',
    '4': '飞镖',
    '5': '礼带',
    '6': '河豚',
    '7': '宝箱',
    '8': '无尽'
}

nameList = []

excelPath = '胧月成绩统计总表.xlsx'

classifyJson = {
    "color": {
        "A": "4EAC5B",
        "B": "A0CD62",
        "C": "FFFD55",
        "D": "F7C143",
        "absence": "EA3323",
        "leave": "BFBFBF"
    },
    "蝙蝠": {
        "A": 390,
        "B": 385,
        "C": 380,
        "max": 401
    },
    "西瓜": {
        "A": 317,
        "B": 315,
        "C": 313,
        "max":317
    },
    "金币": {
        "A": 3090,
        "B": 3080,
        "C": 3070,
        "max":3110
    },
    "飞镖": {
        "A": 600,
        "B": 598,
        "C": 596,
        "max": 601,
    },
    "礼带": {
        "A": 76,
        "B": 73,
        "C": 70,
        "max": 82
    },
    "河豚": {
        "A": 185,
        "B": 180,
        "C": 175,
        "max": 197
    },
    "宝箱": {
        "A": 60,
        "B": 58,
        "C": 57,
        "max": 62
    },
    "无尽": {
        "A": 4067,
        "B": 4066,
        "C": 4065,
        "max": 4067
    }
}

client = AipOcr(**config)

type = input('请选择副本类型 (默认为: 礼带)\n1.蝙蝠 2.西瓜 3.金币 4.飞镖 5.礼带 6.河豚 7.宝箱 8.无尽\n:')
if type == '': type = 5

def get_file_content(file):
    print('识别中...', file)
    try:
        with open(file, 'rb') as fp:
            return fp.read()
    except:
        print('打开文件失败')
        exit(1)

    

def img_to_str(image_path):
    image = get_file_content(image_path)
    result = client.basicGeneral(image)
    data=[]
    if 'words_result' in result:
        for w in result['words_result']:
            data.append(w['words'])
    return data


def getImgInfo():
    print('图片格式必须为: pic+数字，请务必修改')
    photoType = input("请输入图片后缀 (默认为: jpg):")
    if photoType == '': photoType = 'jpg'
    num = input("请输入图片数量 (默认为: 2):")
    if num == '': num = '2'
    count = int(num)
    result = {}
    for i in range(1,count+1):
        imagepath = os.path.dirname(os.path.realpath(sys.executable)) + '/' + 'pic' + str(i) + '.' + photoType
        rawResult = img_to_str(imagepath)
        resultList = []
        for word in rawResult :
            if word.find('新人') < 0 \
                and word.find('族员') < 0 \
                and word.find('精英') < 0 \
                and word.find('豪杰') < 0 \
                and word.find('长老') < 0 \
                and word.find('副族长') < 0 \
                and word.find('族长') <0:
                resultList.append(word)
        j = 0
        while j < len(resultList):
            if j+1 == len(resultList) or not resultList[j+1].isdigit() or resultList[j+1] == '125800':
                maxLength = 0
                if type == '1' or type == '2' or type == '4' or type == '6': #最多3位数字
                    maxLength = 3
                elif type == '3' or type == '8': #最多4位数字
                    maxLength = 4
                else: #其余最多2位数字
                    maxLength = 2
                score = ''
                # print('--------------开始分割成绩--------------')
                # print(resultList[j])
                # print('maxLength:', maxLength)
                for k in range(-1, -maxLength - 1, -1):
                    # print(resultList[j][-1])
                    if resultList[j][-1].isdigit():
                        score = score[::-1]
                        score += resultList[j][-1]
                        score = score[::-1]
                        resultList[j] = resultList[j][:len(resultList[j])-1]
                    else: 
                        score = score[::-1]
                        score += '0'
                        score = score[::-1]
                        break
                resultList.insert(j+1, score)
                # print('--------------分割成绩结束--------------')
            result[resultList[j]] = resultList[j+1]
            j += 2
    print('数据如下')
    print(result)
    return result


def getNameList():
    global excelPath
    inputPath = input('请输入表格名 (默认为: 胧月成绩统计总表.xlsx):')
    if inputPath != '': excelPath = inputPath
    excelPath = os.path.dirname(os.path.realpath(sys.executable)) + '/' + excelPath
    excel = load_workbook(excelPath)
    FuBen = excel['副本']
    nameCells = FuBen['B']
    for cell in nameCells :
        if (cell.value != 'id') and (cell.value != None) and cell.fill.start_color.index != 1 :
            nameList.append(str(cell.value))


def processData(result):
    getNameList()
    col = writeData(result)
    decorateData(col)


def writeData(OCRResult):
    xlsx = load_workbook(excelPath)
    FuBen = xlsx['副本']
    # JiaZuZhan = xlsx['家族战33']
    # ShenYuan = xlsx['深渊']

    col = FuBen.max_column + 1
    inputCol = input('将在表格第%s列插入数据。输入数字更改，或enter确认:' %(col))
    if inputCol != '' : col = int(inputCol)
    maxScore = int(classifyJson[FBType[type]]['max'])
    count = 0
    for name, score in OCRResult.items() :
        row = nameList.index(name) if (name in nameList) else -1
        if row < 0 :
            maxSame = 0
            for i in range(len(nameList) - 1) :
                currentSame = fuzz.partial_ratio(nameList[i], name)
                if currentSame > maxSame :
                    maxSame = currentSame
                    row = i
            if row < 0 :
                print("\033[0;37;41mERROR\033[0m 没有找到族员:\033[0;30;47m%s\033[0m。他的成绩是:\033[0;37;44m%s\033[0m。请确认他是否改名。不然就是程序出错了诶嘿😛" % (name,score))
                continue
            else :
                print("\033[0;30;43mWARN\033[0m 没有找到族员:\033[0;30;47m%s\033[0m，名字最接近的族员是:\033[0;30;47m%s\033[0m。他的成绩是:\033[0;37;44m%s\033[0m。请留意匹配是否出错。" % (name,nameList[row],score))
        
        if int(score) < maxScore * 0.9 :
            print("\033[0;37;40mINFO\033[0m 族员:\033[0;30;47m%s\033[0m的成绩较为异常。他的成绩是:\033[0;37;44m%s\033[0m。请确认是否识别有误。" % (nameList[row],score))
        
        row += 3
        FuBen.cell(row,col).value = int(score)
        count += 1
    
    xlsx.save(excelPath)
    print("\033[0;30;42mSUCCESS\033[0m 成功录入%d条数据！" % (count))
    return col


def decorateData(col):
    xlsx = load_workbook(excelPath)
    FuBen = xlsx['副本']
    classify = classifyJson[FBType[type]]
    colorList = classifyJson['color']
    FuBen.cell(2, col).value = FBType[type]

    border = Border(left=Side(border_style='thin', color='000000'), right=Side(border_style='thin', color='000000'), top=Side(border_style='thin',color='000000'), bottom=Side(border_style='thin',color='000000'))
    
    defaultAbsence = False
    for i in range(3, len(nameList)+ 3): 
        color = ""
        if FuBen.cell(i,col).value == None:
            if not defaultAbsence:
                noScore = input("%s没有成绩，是否请假？(默认否,全部未请假输入all)[N/y/all]:" %(FuBen.cell(i,2).value))
                if noScore == '': noScore = "N"
                while noScore != 'N' and noScore != 'n' and noScore != 'Y' and noScore != 'y' and noScore != 'all':
                    noScore = input("输入无效，请输入‘y’、‘n’、‘all’或直接enter:")
                    if noScore == '': noScore = "N"
                if noScore == 'all': 
                    defaultAbsence = True
                    noScore = 'N'
            else:
                noScore = 'N'       
            color = colorList['absence'] if noScore == "N" or noScore == "n" else colorList['leave']
        elif int(FuBen.cell(i,col).value) >= classify['A']:
            color = colorList['A']
        elif int(FuBen.cell(i,col).value) < classify['A'] and int(FuBen.cell(i,col).value) >= classify['B']:
            color = colorList['B']
        elif  int(FuBen.cell(i,col).value) < classify['B'] and int(FuBen.cell(i,col).value) >= classify['C']:
            color = colorList['C']
        else:
            color = colorList['D']
        FuBen.cell(i, col).fill = PatternFill(fill_type = 'solid', start_color=color, end_color=color)
        FuBen.cell(i, col).border = border
    
    xlsx.save(excelPath)
    print("\033[0;30;42mSUCCESS\033[0m 颜色填充成功！")


if __name__ == '__main__' :
    result = getImgInfo()
    processData(result)
    input("登记完毕。按任意键结束程序……")



