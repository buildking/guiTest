from os import *
import tkinter.messagebox as msbox
import pandas as pd


#엑셀 합치기
def excelCombine():

    #합칠 파일이 부족할 때 경고문
    def combineError(state):
        msbox.showwarning("파일 개수 부족", f"{state}폴더에 합칠 엑셀파일이 부족합니다.")


    # before read
    beforeFileList = listdir('excel_before')

    print(beforeFileList)

    beforeExcelList = []

    for beforeExcel in beforeFileList:
        try:

            beforeExcelThing = pd.DataFrame()
            beforeExcelThing = pd.read_excel(f'./excel_before/{beforeExcel}')

            beforeExcelList.append(beforeExcelThing)

        except:

            msbox.showwarning("파일 인식 불가", "excel_before폴더를 확인해주세요.")

    # after read
    afterFileList = listdir('excel_after')

    afterExcelList = []

    for afterExcel in afterFileList:
        try:
            afterExcelThing = pd.DataFrame()
            afterExcelThing = pd.read_excel(f'./excel_after/{afterExcel}')

            afterExcelList.append(afterExcelThing)

        except:

            msbox.showwarning("파일 인식 불가", "excel_after폴더를 확인해주세요.")



    #before merge

    #폴더 안에 파일이 없거나 하나뿐이라면 에러메세지 띄우기
    if len(beforeExcelList) == 0 or len(beforeExcelList) == 1:
        combineError('before')
        return

    combineBeforeExcel = pd.DataFrame()

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



    #after merge

    #폴더 안에 파일이 없거나 하나뿐이라면 에러메세지 띄우기
    if len(afterExcelList) == 0 or len(afterExcelList) == 1:
        combineError('after')
        return

    combineAfterExcel = pd.DataFrame()

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