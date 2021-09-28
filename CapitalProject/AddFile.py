from tkinter import *
from os import *
import sys
import math
import tkinter.messagebox as msbox
import time
import random
import tkinter.ttk as ttk
from random import *
import threading
import json
from typing import List
import numpy
import pandas as pd
import glob

root = Tk()

########################################파일 읽어오기###################################################

#before read
beforeFileList = listdir('excel_before')

print(beforeFileList)

readExcel = []

for tmp in beforeFileList:

    tmpExcel = pd.read_excel(f'./excel_before/{tmp}')

    readExcel.append(tmpExcel)

print(readExcel)

#after read
beforeFileList2 = listdir('excel_after')

print(beforeFileList2)

readExcel2 = []

for tmp2 in beforeFileList2:

    tmpExcel2 = pd.read_excel(f'./excel_after/{tmp2}')

    readExcel2.append(tmpExcel2)

print(readExcel2)
####################################################################################################################

#에러메시지 출력(프로세스는 조금 더 고민해야될듯)
def combineError():
    pass

#엑셀 합치기
def excelCombine():
    global readExcel

    #before merge
    if len(readExcel) == 0 or len(readExcel) == 1:
        combineError()
        return
    minusexcel = len(readExcel) - 1
    for tmp in range(minusexcel):
        combineExcel = pd.merge(readExcel[tmp], readExcel[tmp+1], how='outer')
    combineExcel.to_excel('./excel_result/excel_before.xlsx', index=False)

    #after merge
    if len(readExcel2) == 0 or len(readExcel2) == 1:
        combineError()
        return
    minusexcel2 = len(readExcel2) - 1
    for tmp in range(minusexcel2):
        combineExcel2 = pd.merge(readExcel2[tmp], readExcel2[tmp+1], how='outer')
    combineExcel2.to_excel('./excel_result/excel_after.xlsx', index=False)


#before, after excel file compare

#dataframe declare
beforeDataFrame = pd.DataFrame()
beforeDataFrameBlood = pd.DataFrame()
beforeDataFrameBloodNum = pd.DataFrame()
beforeDataFrameMom = pd.DataFrame()
beforeDataFrameMomNum = pd.DataFrame()
beforeDataFrameDate = pd.DataFrame()

afterDataFrame = pd.DataFrame()
afterDataFrameBlood = pd.DataFrame()
afterDataFrameBloodNum = pd.DataFrame()
afterDataFrameMom = pd.DataFrame()
afterDataFrameMomNum = pd.DataFrame()
afterDataFrameDate = pd.DataFrame()

beforeCore = pd.DataFrame()
afterCore = pd.DataFrame()

finalResultBefore = pd.DataFrame()
finalResultAfter = pd.DataFrame()

def excelCompare():
    global beforeDataFrame, afterDataFrame

    beforeDataFrame = pd.read_excel('./excel_result/excel_before.xlsx')
    afterDataFrame = pd.read_excel('./excel_result/excel_after.xlsx')

    try :

        #데이터프레임 다이어트 -> 그상태로 엑셀 출력까지?????
        beforeDataFrame = beforeDataFrame[['피보험자', '피보험자 주민등록번호', '모집인명', '모집인 주민번호', '계약체결일']]
        afterDataFrame = afterDataFrame[['피보험자', '피보험자 주민등록번호', '모집인명', '모집인 주민번호', '계약체결일']]

        beforeDataFrameBlood = beforeDataFrame[['피보험자']]
        beforeDataFrameBloodNum = beforeDataFrame[['피보험자 주민등록번호']]
        beforeDataFrameMom = beforeDataFrame[['모집인명']]
        beforeDataFrameMomNum = beforeDataFrame[['모집인명 주민번호']]
        beforeDataFrameDate = beforeDataFrame[['계약체결일']]

        afterDataFrameBlood = afterDataFrame[['피보험자']]
        afterDataFrameBloodNum = afterDataFrame[['피보험자 주민등록번호']]
        afterDataFrameMom = afterDataFrame[['모집인명']]
        afterDataFrameMomNum = afterDataFrame[['모집인명 주민번호']]
        afterDataFrameDate = afterDataFrame[['계약체결일']]

        resultBlood = pd.merge(beforeDataFrameBlood, afterDataFrameBlood, how='inner')
        #주민번호 앞번호 6자리만 따로 빼내기!!
        resultBloodNum = pd.merge(beforeDataFrameBloodNum, afterDataFrameBloodNum, how='inner')
        resultMom = pd.merge(beforeDataFrameMom, afterDataFrameMom, how='inner')
        # 주민번호 앞번호 6자리만 따로 빼내기!!
        resultMomNum = pd.merge(beforeDataFrameMomNum, afterDataFrameMomNum, how='inner')
        #이건 필요없어짐(6개월 차이나기 때문에 비교만 하면 된다!)# resultDate = pd.merge(beforeDataFrameDate, afterDataFrameDate, how='inner')

        for tmp in resultBlood :
            beforeCore.append(beforeDataFrame[beforeDataFrame['피보험자'] == f'{tmp}'])
            afterCore.append(beforeDataFrame[afterDataFrame['피보험자'] == f'{tmp}'])

###########################
#초기 데이터프레임 row data -> 따로 column 만들어서 보관해둬야 한다.
#반드시 번호,데이터 짝지어서 보관해야함!!!!!! ex / (번호,피보험자), (번호,모집인) etc..
#https://rfriend.tistory.com/279?category=675917  <<<<<<<-----------참고!!!!!!!!


        for tmpB in beforeCore :
            for tmpA in afterCore :
                if tmpA['계약체결일'] == tmpB['계약체결일'] and tmpA['모집인명'] == tmpB['모집인명'] and str(tmpA['피보험자 주민등록번호'])[:5] == str(tmpB['피보험자 주민등록번호'])[:5] and str(tmpA['모집인 주민번호'])[:5] == str(tmpB['모집인 주민번호'])[:5]:
                    finalResultBefore.append(tmpB)
                    finalResultAfter.append(tmpA)

        print('%'*180, finalResultBefore,'%'*180, finalResultAfter, '%'*180)





    except : pass

    test1 = pd.DataFrame()
    test1 = pd.merge(beforeDataFrame[['이름', '제품군']], afterDataFrame[['이름', '제품군']], how='inner')
    print('*'*80, beforeDataFrame[['이름','제품군']],'*'*80, afterDataFrame[['이름','제품군']], '*'*180, test1)





#query making
# def queryMake():






root.title("업무 자동화 시스템")
root.geometry("400x300")
root.resizable(False, False)

excelMergeBtn = Button(root, text="엑셀 합치기", command= lambda : excelCombine())
excelMergeBtn.pack(padx=2, pady=2, fill='x')

queryMakeBtn = Button(root, text="쿼리만들기", command=None)
queryMakeBtn.pack(padx=2, pady=2, fill='x')

excelCompareBtn = Button(root, text="엑셀비교하기", command= lambda : excelCompare())
excelCompareBtn.pack(padx=2, pady=2, fill='x')

resultText = Text(root, bg='light gray')
resultText.pack(padx=2, pady=2, fill='both')

root.mainloop()
