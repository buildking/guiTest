from os import *
from tkinter import *
import tkinter.messagebox as msbox
import pandas as pd
import log
import datetime

#log declare
logger = log.setLogging("combine")


#엑셀 합치기
def excelCombine(_text=None):

    #합칠 파일이 부족할 때 경고문
    def combineError(state):
        msbox.showwarning("파일 개수 부족", f"{state}폴더에 합칠 엑셀파일이 부족합니다.")
        logger.debug(f"{state}folder has not enough file")
        print(f"{state}folder has not enough file")

    time = str(datetime.datetime.now())[0:-7]
    _text.insert(END, f"\n[{time}] 엑셀파일 취합을 시작합니다.\n")

    ## before read ##
    #폴더 안의 파일 리스트 생성
    beforeFileList = listdir('excel_before')
    logger.info(beforeFileList)
    print(beforeFileList)
    time = str(datetime.datetime.now())[0:-7]
    _text.insert(END, f"[{time}] 계약 갱신 전 파일 리스트\n{beforeFileList}\n")


    #폴더 안의 파일 저장할 리스트 생성
    beforeExcelList = []
    beforeExcelList.clear()

    #리스트 안의 파일 하나씩 꺼내기
    for beforeExcel in beforeFileList:
        try:

            beforeExcelThing = pd.DataFrame()
            beforeExcelThing = pd.read_excel(f'./excel_before/{beforeExcel}')

            #파일 읽는데 성공했다면, 리스트에 엑셀 저장
            beforeExcelList.append(beforeExcelThing)

        except:

            #엑셀파일 인식이 안된다면, 에러메세지 띄우기
            msbox.showwarning("파일 인식 불가", "excel_before폴더를 확인해주세요.")
            logger.debug("before excel file don't observed")
            print("before excel file don't observed")

    ## after read ##
    #폴더 안의 파일 리스트 생성
    afterFileList = listdir('excel_after')
    logger.info(afterFileList)
    print(afterFileList)
    time = str(datetime.datetime.now())[0:-7]
    _text.insert(END, f"[{time}] 계약 갱신 후 파일 리스트\n{afterFileList}\n")

    #폴더 안의 파일 저장할 리스트 생성
    afterExcelList = []
    afterExcelList.clear()

    #파일리스트 안의 파일 하나씩 꺼내기
    for afterExcel in afterFileList:
        try:
            afterExcelThing = pd.DataFrame()
            afterExcelThing = pd.read_excel(f'./excel_after/{afterExcel}')

            #파일 읽는데 성공했다면, 리스트에 엑셀 저장
            afterExcelList.append(afterExcelThing)

        except:

            #엑셀파일 인식이 안된다면, 에러메세지 띄우기
            msbox.showwarning("파일 인식 불가", "excel_after폴더를 확인해주세요.")
            logger.debug("after excel file don't observed")
            print("after excel file don't observed")

    #before merge

    #폴더 안에 파일이 없거나 하나뿐이라면 에러메세지 띄우기
    if len(beforeExcelList) == 0 or len(beforeExcelList) == 1:
        combineError('before')
        return


    #첫번째 엑셀파일을 데이터프레임에 삽입
    combineBeforeExcel = beforeExcelList[0]

    #폴더 안에 있는 엑셀 리스트들을 차례대로 병합
    for beforeExcelOne in beforeExcelList[1:]:
        combineBeforeExcel = pd.concat([beforeExcelOne, combineBeforeExcel], join='outer')

    # 만들어진 엑셀 행*열 개수 tuple형태로 저장
    combineBeforeExcelShape = combineBeforeExcel.shape

    #기존 '연번' 컬럼 삭제
    combineBeforeExcel = combineBeforeExcel.drop(columns='연번', axis=1)

    #새로운 '연번'에 들어갈 list 생성
    beForeIndexList = []
    beForeIndexList.clear()

    #엑셀 row 수 세기
    x = 0
    for index in range(combineBeforeExcelShape[0]):
        x += 1
        beForeIndexList.append(x)

    #새로 생성한 '연번'을 첫번째 칼럼으로 삽입
    combineBeforeExcel.insert(0, '연번', beForeIndexList)
    #before파일 엑셀로 만들기
    combineBeforeExcel.to_excel('./excel_result/excel_before.xlsx', index=False)
    time = str(datetime.datetime.now())[0:-7]
    _text.insert(END, f"[{time}] 계약 갱신 전 엑셀파일 취합에 성공했습니다. \n")


    #after merge

    #폴더 안에 파일이 없거나 하나뿐이라면 에러메세지 띄우기
    if len(afterExcelList) == 0 or len(afterExcelList) == 1:
        combineError('after')
        return


    #첫번째 엑셀파일을 데이터프레임에 삽입
    combineAfterExcel = afterExcelList[0]

    #폴더 안에 있는 엑셀 리스트들을 차례대로 병합
    for afterExcelOne in afterExcelList[1:]:
        combineAfterExcel = pd.concat([afterExcelOne, combineAfterExcel], join='outer')

    # 만들어진 엑셀 행*열 개수 tuple형태로 저장
    combineAfterExcelShape = combineAfterExcel.shape

    #기존 '연번' 컬럼 삭제
    combineAfterExcel = combineAfterExcel.drop(columns='연번', axis=1)


    #새로운 '연번'에 들어갈 list 생성
    afterIndexList = []
    afterIndexList.clear()

    #엑셀 row 수 세기
    x = 0
    for index in range(combineAfterExcelShape[0]):
        x += 1
        afterIndexList.append(x)

    #새로 생성한 '연번'을 첫번째 칼럼으로 삽입
    combineAfterExcel.insert(0, '연번', afterIndexList)
    #after파일 엑셀로 만들기
    combineAfterExcel.to_excel('./excel_result/excel_after.xlsx', index=False)
    time = str(datetime.datetime.now())[0:-7]
    _text.insert(END, f"[{time}] 계약 갱신 후 엑셀파일 취합에 성공했습니다. \n")
    time = str(datetime.datetime.now())[0:-7]
    _text.insert(END, f"[{time}] 엑셀파일 취합을 종료합니다.\n")
    _text.see(END)
    logger.info("combine complete")
    print("Excel file combine complete")
