from openpyxl.styles import PatternFill
import openpyxl
from tkinter import *
import pandas as pd
import wemade_log
import datetime
import time as Time
from os import *
from wemade_main import *
import numpy as np

logger = wemade_log.setLogging("compare")

#end, new excel file compare
def xlCompare(_text=None):

    time = str(datetime.datetime.now())[0:-7]
    _text.insert(END, f"\n[{time}] 엑셀파일을 불러옵니다.\n")

    #폴더 안 파일의 리스트 생성
    excelList = listdir('input')
    logger.info(excelList)
    print(excelList)
    time = str(datetime.datetime.now())[0:-7]
    _text.insert(END, f"[{time}] 파일 리스트\n{excelList}\n")

    #폴더 안에 파일이 2개가 아니라면 에러메세지 띄우기
    if len(excelList) != 2:
        shortError('input')
        return




    xl = pd.ExcelFile(f'./input/{excelList[1]}')
    #엑셀 sheet 개수
    xlSheet = len(xl.sheet_names)
    print('시트 개수 = ',xlSheet)
    #엑셀 sheet별 column*row
    xlShape = []


    # 데이터가 다른 cell의 column, row 저장(append로 추가)
    coreRowNum = []
    coreRowNum.clear()
    coreRowNumPrint = []
    coreRowNumPrint.clear()

    for sheetNum in range(xlSheet):
        firExcel = pd.read_excel(f'./input/{excelList[0]}', sheet_name = sheetNum, dtype = str)
        secExcel = pd.read_excel(f'./input/{excelList[1]}', sheet_name = sheetNum, dtype = str)

        #데이터프레임 >> 리스트
        f_list = firExcel.values.tolist()
        s_list = secExcel.values.tolist()

        #sheet별 column, row
        xlRow = len(secExcel.index)
        xlColumn = len(secExcel.columns)

        xlShape.append([xlRow, xlColumn])
        print(sheetNum, '번째 시트 크기 = ',xlRow, 'X', xlColumn)


        time = str(datetime.datetime.now())[0:-7]
        _text.insert(END, f"\n[{time}] {sheetNum}번째 시트 데이터 변경점 탐색을 시작합니다.\n")

        start = Time.time()

        i = 0
        j = 0
        for _ in range(xlRow):
            j = 0
            for _ in range(xlColumn):
                ######비교 로직 작성

                if str(f_list[i][j]) != str(s_list[i][j]):
                    coreRowNum.append([sheetNum, i, j])
                    coreRowNumPrint.append([sheetNum, i + 1, j + 1])

                ######
                j = j + 1
            i = i + 1




    end = Time.time()
    logger.info("excute time >>>>> " + str(end - start) + "s")

    logger.debug("변경된 데이터 리스트")
    logger.debug(coreRowNumPrint)
    logger.debug(len(coreRowNumPrint))
    print("변경된 데이터 리스트")
    print(coreRowNumPrint)
    print(len(coreRowNumPrint))

    #데이터 이격 생긴 부분 컬럼, 로우만 갈무리 해서 엑셀 생성
    coreRowNumExcel = pd.DataFrame(coreRowNumPrint, columns=['sheet','row','column'])
    coreRowNumExcel.to_excel('./output/excel_tryout_rows.xlsx')

    #제목에 넣을 날짜
    finishStamp = str(datetime.datetime.now())[0:10]
    #cell에 색칠할 엑셀 output에 저장
    xlTemp = openpyxl.load_workbook(filename = f'./input/{excelList[1]}')
    xlTemp.save(filename = f'./output/{excelList[0],excelList[1],finishStamp}.xlsx')
    xlTemp.close()


    #cell 색칠 시작
    time = str(datetime.datetime.now())[0:-7]
    _text.insert(END, f"[{time}] painting start \n")
    logger.info("painting start")
    print("Excel file painting start")

    #노란색 선언
    y_color = PatternFill(start_color='ffff99', end_color='ffff99', fill_type='solid')

    #openpyxl을 통해 엑셀시트를 연다.
    pyxl = openpyxl.load_workbook(filename=f'./output/{excelList[0],excelList[1],finishStamp}.xlsx')

    for cellSheet, cellRow, cellCol in coreRowNumPrint:

        pyxlSheet = pyxl.worksheets[cellSheet]

        #i=column, j=row
        pyxlSheet.cell(cellRow+1, cellCol).fill = y_color

    pyxl.save(filename=f'./output/{excelList[0],excelList[1],finishStamp}.xlsx')

    pyxl.close()


    time = str(datetime.datetime.now())[0:-7]
    _text.insert(END, f"[{time}] painting complete")
    logger.info("painting complete")
    print("Excel file painting complete")

    time = str(datetime.datetime.now())[0:-7]
    _text.insert(END, f"[{time}] 데이터 변경점 검색을 종료합니다.\n")
    _text.see(END)