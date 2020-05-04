# -*- coding:UTF-8 -*-
# -*- encoding: utf-8 -*-
from abyss import abyss
from dungeon import dungeonPreparation, familyWarDungeon
from fight import familyWarFight
from utils import refreshAccessToken

if __name__ == '__main__' :
    refreshAccessToken()
    job = input("请选择要登记的成绩\n1.备战副本\n2.深渊\n3.家族战副本\n4.家族战33\n(1/2/3/4):")
    while 1:
        if job == '1':
            dungeonPreparation()
            break
        elif job == '2':
            abyss()
            break
        elif job == '3':
            familyWarDungeon()
            break
        elif job == '4':
            familyWarFight()
            break
        else:
            print('不合法输入，请重新输入。')
            job = input("请选择要登记的成绩\n1.备战副本\n2.深渊\n3.家族战副本\n4.家族战33\n(1/2/3/4):")
    input("登记完毕。按任意键结束程序……")
