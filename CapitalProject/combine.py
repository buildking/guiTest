from os import *
import tkinter.messagebox as msbox
import pandas as pd


#엑셀 합치기
def excelCombine():

    #합칠 파일이 부족할 때 경고문
    def combineError():
        msbox.showwarning("파일 개수 부족", "합칠 엑셀파일이 부족합니다.")



#####################################파일 읽어오기##################################

    # before read
    beforeFileList = listdir('excel_before')

    print(beforeFileList)

    readExcel = []

    for tmp in beforeFileList:
        tmpExcel = pd.read_excel(f'./excel_before/{tmp}')

        readExcel.append(tmpExcel)

    # after read
    beforeFileList2 = listdir('excel_after')

    readExcel2 = []

    for tmp2 in beforeFileList2:
        tmpExcel2 = pd.read_excel(f'./excel_after/{tmp2}')

        readExcel2.append(tmpExcel2)

##########################################################################

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

