from tkinter import *
import pandas as pd
import log
import datetime

logger = log.setLogging("compare")

#before, after excel file compare
def excelCompare(_text=None):

    time = str(datetime.datetime.now())[0:-7]
    _text.insert(END, f"\n[{time}] 계약 갱신 전,후 공통계약 찾기를 시작합니다.\n")

    beforeDataFrame = pd.read_excel('./excel_result/excel_before.xlsx', dtype=str)
    afterDataFrame = pd.read_excel('./excel_result/excel_after.xlsx', dtype=str)
    print('um')

    #비교 통과한 before data와 after data의 row 저장(append로)
    coreRowNum = []
    coreRowNum.clear()
    coreRowNumPrint = []
    coreRowNumPrint.clear()

    for pdb, rowBefore in beforeDataFrame.iterrows():
        for pda, rowAfter in afterDataFrame.iterrows():
            if str(rowBefore['피보험자']) == str(rowAfter['피보험자']):
                if str(rowBefore['모집인명']) == str(rowAfter['모집인명']):
                    if str(rowBefore['피보험자\n주민등록번호'])[:5] == str(rowAfter['피보험자\n주민등록번호'])[:5]:
                        if str(rowBefore['모집인 주민번호'])[:5] == str(rowAfter['모집인 주민번호'])[:5]:
                            coreRowNum.append([pdb, pda])
                            coreRowNumPrint.append([pdb+1, pda+1])

    print(coreRowNumPrint)
    logger.info("계약일 제외하고 일치하는 리스트")
    logger.info(coreRowNumPrint)
    coreRowNumPrint = pd.DataFrame(coreRowNumPrint, columns=['전','후'])
    coreRowNumPrint.to_excel('./excel_result/excel_final_tryout_rows.xlsx')

    #1차로 걸러진 데이터 출력
    finalExcel_tryout = pd.DataFrame(columns=[['소멸계약<이동 전>', '소멸계약<이동 전>', '소멸계약<이동 전>', '소멸계약<이동 전>', '소멸계약<이동 전>', '소멸계약<이동 전>', '소멸계약<이동 전>', '신규계약<이동 후>', '신규계약<이동 후>', '신규계약<이동 후>', '신규계약<이동 후>', '신규계약<이동 후>', '신규계약<이동 후>', '신규계약<이동 후>', '신규계약<이동 후>', '신규계약<이동 후>'],['회사명', '피보험자', '증권번호', '상품명', '상품분류', '계약소멸일', '상태', '회사명', '피보험자', '증권번호', '상품명', '상품분류', '계약체결일', '상태', '모집인명', '생년월일']])

    x = 0
    for b, a in coreRowNum:
        x += 1
        finalExcel_tryout.loc[x] = (beforeDataFrame.loc[b, '회사명'], beforeDataFrame.loc[b, '피보험자'], beforeDataFrame.loc[b, '증권번호'], beforeDataFrame.loc[b, '상품명'], beforeDataFrame.loc[b, '상품분류'], beforeDataFrame.loc[b, '계약소멸일'], beforeDataFrame.loc[b, '계약상태'], afterDataFrame.loc[a, '회사명'], afterDataFrame.loc[a, '피보험자'], afterDataFrame.loc[a, '증권번호'], afterDataFrame.loc[a, '상품명'], afterDataFrame.loc[a, '상품분류'], afterDataFrame.loc[a, '계약체결일'], afterDataFrame.loc[a, '계약상태'], afterDataFrame.loc[a, '모집인명'], afterDataFrame.loc[a, '피보험자\n주민등록번호'])

    #1차로 걸러진 데이터 엑셀로 생성
    finalExcel_tryout.to_excel('./excel_result/excel_final_tryout.xlsx')


    #lastList = before[계약소멸일] - after[계약체결일] 의 list
    lastList = []
    lastList.clear()
    #lastListPrint = 사용자가 보기 편하게 출력용으로 제작
    lastListPrint = []
    lastListPrint.clear()

    #날짜 분류 함수화
    def dateDivision(date):
        #변수의 길이 측정
        dateLen = len(date)
        #str 형식으로, 하이픈 없이 입력됐다면 6글자 미만일것이라고 가정한다. 혹시 몰라서 7글자를 기준으로 잡음.
        if dateLen < 7 :
            return int(date)
        else:
            dateResult = ((int(date[0:3])-1900)*365) + (int(date[5:6])*30) + int(date[8:9])
            return dateResult

    #조건을 통과한 값의 (b'계약소멸일' - a'계약체결일')이 180이내인지 추려냄
    for i, j in coreRowNum:
        # beforeDate = len(beforeDataFrame.loc[i, '계약소멸일'])
        # afterDate = len(afterDataFrame.loc[j, '계약체결일'])
        beforeDateNum = dateDivision(beforeDataFrame.loc[i, '계약소멸일'])
        afterDateNum = dateDivision(afterDataFrame.loc[j, '계약체결일'])
        if abs(afterDateNum-beforeDateNum) <= 180:
            aSubB = afterDateNum-beforeDateNum
            lastList.append([i, j, aSubB])
            lastListPrint.append([i+1, j+1, aSubB])

    #개발자용
    lastExcel = pd.DataFrame(lastList, columns=['전','후', '날짜 차이'])
    logger.info(lastExcel)
    print(lastExcel)
    #사용자용
    lastExcelPrint = pd.DataFrame(lastListPrint, columns=['전','후', '날짜 차이'])
    lastExcelPrint.to_excel('./excel_result/excel_final_rows.xlsx', index=False)
    time = str(datetime.datetime.now())[0:-7]
    _text.insert(END, f"[{time}] 확인된 공통계약 리스트\n{lastExcelPrint}\n")

    beforeFinal = pd.DataFrame
    afterFinal = pd.DataFrame

    beforeFinal = pd.concat([beforeDataFrame.loc[[b]] for b, a, diff in lastList], ignore_index=True)
    afterFinal = pd.concat([afterDataFrame.loc[[a]] for b, a, diff in lastList], ignore_index=True)

    columnList = beforeDataFrame.columns.tolist()
    for afcol in afterDataFrame.columns.tolist():
        columnList.append(afcol)
    columnList.append('날짜 차이')

    print(columnList, len(columnList))

    finalExcel = pd.concat([beforeFinal, afterFinal, lastExcel.iloc[:, 2]], axis=1, ignore_index=True)
    finalExcel.columns = columnList

    finalExcel.to_excel('./excel_result/excel_final.xlsx')
    logger.info("compare complete")
    print("Excel file compare complete")
    time = str(datetime.datetime.now())[0:-7]
    _text.insert(END, f"[{time}] 공통계약 찾기를 종료합니다.\n")
    _text.see(END)