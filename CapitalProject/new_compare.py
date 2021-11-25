from tkinter import *
import pandas as pd
import log
import datetime
import time as Time
from DbUtil import DbUtil


#로그파일 세팅
logger = log.setLogging("queryMake_sqlite")
logger.debug("config file read")

dbUtil = DbUtil('db/knia.db')#손해보험협회 약어(knia)

def dbcompare(_text=None):

    time = str(datetime.datetime.now())[0:-7]
    _text.insert(END, f"\n[{time}] 계약 갱신 전,후 공통계약 찾기를 시작합니다.\n")

    dbRow = dbUtil.selectCompareResult()
    resultExcel = pd.DataFrame(dbRow)
    print(resultExcel)
    resultExcel.to_excel('./excel_result/excel_final_result.xlsx')

    logger.info("DB 비교 완료")

    time = str(datetime.datetime.now())[0:-7]
    _text.insert(END, f"[{time}] 공통계약 찾기를 종료합니다.\n")
    _text.see(END)