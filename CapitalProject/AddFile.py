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
import numpy as np
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

#print(readExcel)

#after read
beforeFileList2 = listdir('excel_after')

#print(beforeFileList2)

readExcel2 = []

for tmp2 in beforeFileList2:

    tmpExcel2 = pd.read_excel(f'./excel_after/{tmp2}')

    readExcel2.append(tmpExcel2)

#print(readExcel2)
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

    # print(beforeDataFrame)
    # print(afterDataFrame)

    #############################pandasVersion#####################################
    #
    # #데이터프레임 다이어트 -> 그상태로 엑셀 출력까지?????
    # beforeDataFrame = beforeDataFrame[['피보험자', '피보험자 주민등록번호', '모집인명', '모집인 주민번호', '계약체결일']]
    # beforeDataFrame['beforeRow'] = beforeDataFrame['index']
    # afterDataFrame = afterDataFrame[['피보험자', '피보험자 주민등록번호', '모집인명', '모집인 주민번호', '계약체결일']]
    # afterDataFrame['afterRow'] = afterDataFrame['index']
    #
    # beforeDataFrameBlood = beforeDataFrame[['피보험자', 'beforeRow']]
    # beforeDataFrameBloodNum = beforeDataFrame[['피보험자 주민등록번호', 'beforeRow']]
    # beforeDataFrameMom = beforeDataFrame[['모집인명', 'beforeRow']]
    # beforeDataFrameMomNum = beforeDataFrame[['모집인명 주민번호', 'beforeRow']]
    # beforeDataFrameDate = beforeDataFrame[['계약체결일', 'beforeRow']]
    #
    # afterDataFrameBlood = afterDataFrame[['피보험자', 'afterRow']]
    # afterDataFrameBloodNum = afterDataFrame[['피보험자 주민등록번호', 'afterRow']]
    # afterDataFrameMom = afterDataFrame[['모집인명', 'afterRow']]
    # afterDataFrameMomNum = afterDataFrame[['모집인명 주민번호', 'afterRow']]
    # afterDataFrameDate = afterDataFrame[['계약체결일', 'afterRow']]
    #
    # resultBlood = pd.merge(beforeDataFrameBlood, afterDataFrameBlood, how='inner')
    # #주민번호 앞번호 6자리만 따로 빼내기!!
    # resultBloodNum = pd.merge(beforeDataFrameBloodNum, afterDataFrameBloodNum, how='inner')
    # resultMom = pd.merge(beforeDataFrameMom, afterDataFrameMom, how='inner')
    # # 주민번호 앞번호 6자리만 따로 빼내기!!
    # resultMomNum = pd.merge(beforeDataFrameMomNum, afterDataFrameMomNum, how='inner')
    # #이건 필요없어짐(6개월 차이나기 때문에 비교만 하면 된다!)# resultDate = pd.merge(beforeDataFrameDate, afterDataFrameDate, how='inner')
    #
    # for tmp in resultBlood :
    #     beforeCore.append(beforeDataFrame[beforeDataFrame['피보험자'] == f'{tmp}'])
    #     afterCore.append(afterDataFrame[afterDataFrame['피보험자'] == f'{tmp}'])

###########################
#초기 데이터프레임 row data -> 따로 column 만들어서 보관해둬야 한다.
#반드시 번호,데이터 짝지어서 보관해야함!!!!!! ex / (번호,피보험자), (번호,모집인) etc..
#https://rfriend.tistory.com/279?category=675917  <<<<<<<-----------참고!!!!!!!!
    # for tmpB in beforeCore :
    #     for tmpA in afterCore :
    #         if tmpA['계약체결일'] == tmpB['계약체결일'] and tmpA['모집인명'] == tmpB['모집인명'] and str(tmpA['피보험자 주민등록번호'])[:5] == str(tmpB['피보험자 주민등록번호'])[:5] and str(tmpA['모집인 주민번호'])[:5] == str(tmpB['모집인 주민번호'])[:5]:
    #             finalResultBefore.append(tmpB)
    #             finalResultAfter.append(tmpA)
    # print('%'*180, finalResultBefore,'%'*180, finalResultAfter, '%'*180)
########################################################################################


    #비교 통과한 before data와 after data의 row 저장(append로)
    coreRowNum = []
    coreRowNum.clear()


    for pdb, rowBefore in beforeDataFrame.iterrows():
        print('되냐??'*10)
        print(pdb,'뎃데로게'*10, rowBefore)
        print('야지막으로, ', rowBefore['피보험자\n주민등록번호'])
##,'모집인명','모집인 주민번호','계약체결일']
# #####################################무지성 for문 때려박기########################
#
#     #before data row
#     beforeRowNum = 0
#     #after data row
#     afterRowNum = 0
#
#     #비포엑셀 데이터 한줄씩 출력
#     for bdf in beforeDataFrame:
#         #몇번째 비포데이터인지 입력할 예정
#         beforeRowNum += 1
#         #애프터엑셀 데이터 한줄씩 출력
#         for adf in afterDataFrame:
#             #몇번째 애프터데이터인지 입력할 예정
#             afterRowNum += 1
#             #비교 조건 입력
#             if beforeDataFrame.loc[f'{beforeRowNum-1}','피보험자'] == afterDataFrame.loc[f'{afterRowNum-1}','피보험자']:
#                 if beforeDataFrame.loc[f'{beforeRowNum-1}','모집인명'] == afterDataFrame.loc[f'{afterRowNum-1}','모집인명']:
#                     if str(beforeDataFrame.loc[f'{beforeRowNum-1}','피보험자\n주민등록번호'])[:5] == str(afterDataFrame.loc[f'{afterRowNum-1}','피보험자\n주민등록번호'])[:5]:
#                         if str(beforeDataFrame.loc[f'{beforeRowNum-1}','모집인 주민번호'])[:5] == str(afterDataFrame.loc[f'{afterRowNum-1}','모집인 주민번호'])[:5]:
#                             coreRowNum.append([beforeRowNum-1, afterRowNum-1])
# # #####################################################################################

#
########################판다스 모듈로 ㄱㄱ############################################
    for pdb, rowBefore in beforeDataFrame.iterrows():
        print(pdb, rowBefore)
        for pda, rowAfter in afterDataFrame.iterrows():
            print(pda, rowAfter)
            if str(rowBefore['피보험자']) == str(rowAfter['피보험자']):
                if str(rowBefore['모집인명']) == str(rowAfter['모집인명']):
                    if str(rowBefore['피보험자\n주민등록번호'])[:5] == str(rowAfter['피보험자\n주민등록번호'])[:5]:
                        if str(rowBefore['모집인 주민번호'])[:5] == str(rowAfter['모집인 주민번호'])[:5]:
                            coreRowNum.append([pdb, pda])
###################################################################################
#

    print('*'*50, coreRowNum, '*'*50)

    lastnum = []
    lastnum.clear()
    for i, j in coreRowNum:
        print('되나연?????'*10)
        print(beforeDataFrame.loc[i,'계약체결일'])
        print(type(beforeDataFrame.loc[i, '계약체결일']))
        print(afterDataFrame.loc[j,'계약체결일'])
        print(type(afterDataFrame.loc[j,'계약체결일']))
        print(i, j)
        print(afterDataFrame.loc[j,'계약체결일']-beforeDataFrame.loc[i,'계약체결일'])
        print(str(afterDataFrame.loc[j,'계약체결일']-beforeDataFrame.loc[i,'계약체결일']))
        print(str(afterDataFrame.loc[j,'계약체결일']-beforeDataFrame.loc[i,'계약체결일']).split()[0])
        #181일로 할지 180일로 할지 물어봐야됨. -> 181 확정

        if int(str(pd.to_datetime(str(afterDataFrame.loc[j,'계약체결일']))-pd.to_datetime(str(beforeDataFrame.loc[i,'계약체결일']))).split()[0]) < 181 :
            lastnum.append([i,j])

    print(lastnum)





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
