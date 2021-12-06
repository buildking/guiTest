import openpyxl
from tkinter import *
import pandas as pd
import log
import datetime
import time as Time

logger = log.setLogging("compare")

#end, new excel file compare
def xlCompare(_text=None):

    #합칠 파일이 부족할 때 경고문
    def shortError(state):
        msbox.showwarning("파일 개수 오류", f"{state}폴더에는 엑셀파일을 반드시 2개 넣어야 합니다.")
        logger.debug(f"{state}folder's file must be 2")
        print(f"{state}folder's file must be 2")

    time = str(datetime.datetime.now())[0:-7]
    _text.insert(END, f"\n[{time}] 엑셀파일을 불러옵니다.\n")

    #폴더 안 파일의 리스트 생성
    FileList = listdir('input')
    logger.info(FileList)
    print(FileList)
    time = str(datetime.datetime.now())[0:-7]
    _text.insert(END, f"[{time}] 파일 리스트\n{FileList}\n")

    #폴더 안에 파일이 2개가 아니라면 에러메세지 띄우기
    if len(FileList) != 2:
        shortError('input')
        return

    #리스트 안의 파일 하나씩 꺼내기
    #첫번째 파일
    try:
        endexcel = pd.DataFrame()
        #엑셀 모든 시트를 합쳐서 받는다.
        endexcel = pd.read_excel(f'./input/{FileList[0]}', sheet_name=None, dtype=str)
    except:

        #엑셀파일 인식이 안된다면, 에러메세지 띄우기
        msbox.showwarning("파일 인식 불가", f"input폴더의 {FileList[0]}파일을 확인해주세요.")
        logger.debug("excel file don't observed")
        print("excel file don't observed")
    #두번째 파일
    try:
        newexcel = pd.DataFrame()
        #엑셀 모든 시트를 합쳐서 받는다.
        newexcel = pd.read_excel(f'./input/{FileList[1]}', sheet_name=None, dtype=str)

    except:
        #엑셀파일 인식이 안된다면, 에러메세지 띄우기
        msbox.showwarning("파일 인식 불가", f"input폴더의 {FileList[1]}파일을 확인해주세요.")
        logger.debug("excel file don't observed")
        print("excel file don't observed")




#######비교

    time = str(datetime.datetime.now())[0:-7]
    _text.insert(END, f"\n[{time}] 데이터 변경점 검색을 시작합니다.\n")

    start = Time.time()


#################공구리예정############################
    #dataframe을 dictionary type으로 변환
    b_dict = endexcel.to_dict("records")
    a_dict = newexcel.to_dict("records")
    i = 0
    j = 0
    for b in b_dict:
        j = 0
        for a in a_dict:
            ######비교 로직 작성
            if str(b['피보험자']) == str(a['피보험자']):
                if str(b['모집인명']) == str(a['모집인명']):
                    if str(b['피보험자\n주민등록번호'])[:5] == str(a['피보험자\n주민등록번호'])[:5]:
                        if str(b['모집인 주민번호'])[:5] == str(a['모집인 주민번호'])[:5]:
                            coreRowNum.append([i, j])
                            coreRowNumPrint.append([i+1, j+1])
            ######
            j = j + 1
        i = i + 1

    end = Time.time()
    logger.info("excute time >>>>> " + str(end - start) + "s")

    logger.debug("계약일 제외하고 일치하는 리스트")
    logger.debug(coreRowNumPrint)
    logger.debug(len(coreRowNumPrint))

    coreRowNumPrint = pd.DataFrame(coreRowNumPrint, columns=['전','후'])
    coreRowNumPrint.to_excel('./output/excel_tryout_rows.xlsx')

    #개발자용
    lastExcel = pd.DataFrame(lastList, columns=['전','후', '날짜 차이'])
    logger.debug(lastExcel)#개발자용이기때문에 log에 남길필요 없음
    #사용자용
    lastExcelPrint = pd.DataFrame(lastListPrint, columns=['전','후', '날짜 차이'])
    lastExcelPrint.to_excel('./excel_result/excel_final_rows.xlsx', index=False)
    ###############공구리예정###################



    time = str(datetime.datetime.now())[0:-7]
    _text.insert(END, f"[{time}] 확인된 공통계약 리스트\n{lastExcelPrint}\n")

    finishStamp = str(datetime.datetime.now())[0:10]
    finalExcel.to_excel(f'./output/{FileList[0],FileList[1]-finishStamp}.xlsx')
    logger.info("compare complete")
    print("Excel file compare complete")
    time = str(datetime.datetime.now())[0:-7]
    _text.insert(END, f"[{time}] 데이터 변경점 검색을 종료합니다.\n")
    _text.see(END)