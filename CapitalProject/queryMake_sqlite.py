import pandas as pd
import log
import datetime
from tkinter import *
import threading
from DbUtil import DbUtil

#로그파일 세팅
logger = log.setLogging("queryMake_sqlite")
logger.debug("config file read")

dbUtil = DbUtil('db/knia.db')#손해보험협회 약어(knia)

def queryMake(_textEntry=None):
    logger.info("query make start")

    newContractInsert()
    endContractInsert()

    time = str(datetime.datetime.now())[0:-7]
    _textEntry.insert(END, "[{}] {}".format(time, '정상적으로 저장되었습니다.\n'))
    _textEntry.see(END)
    dbUtil.databaseClose()

def endContractInsert():

    #DataFrame을 가져온다.
    beforeDf = pd.read_excel('./excel_result/excel_before.xlsx')

    logger.info("before_dataFrame is ready")

    _paramList = []
    for index, row in beforeDf.iterrows():#beforeDf의 row를 순회
        # N번쨰 row 가져오기
        # 각 컬럼이 16개나 되기때문에 index로 접근하기위해 .loc사용
        col0 = str(beforeDf.loc[index][0])#연번
        col4 = str(beforeDf.loc[index][4]).strip()#피보험자
        col5 = str(beforeDf.loc[index][5]).strip()[:6]#피보험 주민번호
        #col10 = str(beforeDf.loc[index][10]).strip().split(" ")[0]#계약체결일
        col11 = str(beforeDf.loc[index][9]).strip().split(" ")[0]#계약소멸일
        col13 = str(beforeDf.loc[index][11]).strip()#모집인명
        col14 = str(beforeDf.loc[index][12]).strip()[:6]#모집인주민번호

        _paramList.append({
            'numberNew': col0,
            'newDate': col11,
            'insNm': col4,
            'insBirth': col5,
            'plnrNm': col13,
            'plnrBirth': col14,
        })
    dbUtil.deleteEndContract()
    dbUtil.insertEndContract(_paramList)
    logger.info("before_excel to sql success")
    #file.close()

    # TODO 현재까지 파악된 바에 따르면
    # t_before.join()으로 t_before가 끝날때까지 대기하는데
    # _textEntry.insert에서 Thread를 점유하여 처리해야 하므로 대기하게된다.
    # 그래서 t_before가 끝나지 않고 deadlock이 걸린다.
    #time = str(datetime.datetime.now())[0:-7]
    #_textEntry.insert(END, "[{}] {}".format(time, 'before 쿼리만들기가 완료되었습니다.\n'))

def newContractInsert():

    #afterDataFrame을 가져온다.
    afterDf = pd.read_excel('./excel_result/excel_after.xlsx')

    logger.info("after_dataFrame is ready")

    _paramList = []
    for index, row in afterDf.iterrows():#afterDf의 row를 순회
        # N번쨰 row 가져오기
        # 각 컬럼이 16개나 되기때문에 index로 접근하기위해 .loc사용
        col0 = str(afterDf.loc[index][0])  # 연번
        col4 = str(afterDf.loc[index][4]).strip()  # 피보험자
        col5 = str(afterDf.loc[index][5]).strip()[:6]  # 피보험 주민번호
        col10 = str(afterDf.loc[index][9]).strip().split(" ")[0]  # 계약체결일
        #col11 = str(afterDf.loc[index][11]).strip().split(" ")[0]  # 계약소멸일
        col13 = str(afterDf.loc[index][11]).strip()  # 모집인명
        col14 = str(afterDf.loc[index][12]).strip()[:6]  # 모집인주민번호

        _paramList.append({
            'numberEnd': col0,
            'endDate': col10,
            'insNm': col4,
            'insBirth': col5,
            'plnrNm': col13,
            'plnrBirth': col14,
        })

    dbUtil.deleteNewContract()
    dbUtil.insertNewContract(_paramList)
    logger.info("after_excel to sql success")