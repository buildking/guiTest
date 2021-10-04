from os import *
import tkinter.messagebox as msbox
import pandas as pd


#엑셀 합치기
def excelCombine():

    #합칠 파일이 부족할 때 경고문
    def combineError(state):
        msbox.showwarning("파일 개수 부족", f"{state}폴더에 합칠 엑셀파일이 부족합니다.")



#####################################파일 읽어오기##################################

    # before read
    beforeFileList = listdir('excel_before')

    print(beforeFileList)

    beforeExcelList = []

    for beforeExcel in beforeFileList:
        try:

            beforeExcelThing = pd.read_excel(f'./excel_before/{beforeExcel}')

            beforeExcelList.append(beforeExcelThing)

        except:

            msbox.showwarning("파일 인식 불가", "excel_before폴더를 확인해주세요.")

    # after read
    afterFileList = listdir('excel_after')

    afterExcelList = []

    for afterExcel in afterFileList:
        try:
            afterExcelThing = pd.read_excel(f'./excel_after/{afterExcel}')

            afterExcelList.append(afterExcelThing)

        except:

            msbox.showwarning("파일 인식 불가", "excel_after폴더를 확인해주세요.")

    ##########################################################################

    #before merge
    if len(beforeExcelList) == 0 or len(beforeExcelList) == 1:
        combineError('before')
        return
    combineBeforeExcel = pd.DataFrame()
    for beforeExcelOne in range(beforeExcelList):
        combineBeforeExcel = pd.merge(combineBeforeExcel, beforeExcelOne, how='outer')
    combineBeforeExcel.to_excel('./excel_result/excel_before.xlsx', index=False)

    #after merge
    if len(afterExcelList) == 0 or len(afterExcelList) == 1:
        combineError('after')
        return
    combineAfterExcel = pd.DataFrame()
    for afterExcelOne in range(afterExcelList):
        combineAfterExcel = pd.merge(combineAfterExcel, afterExcelOne, how='outer')
    combineAfterExcel.to_excel('./excel_result/excel_after.xlsx', index=False)

