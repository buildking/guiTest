import pandas as pd
import cx_Oracle
import log
import os
from zipfile import ZipFile
from config import ConfigUtil
import datetime
from tkinter import *
import threading

#로그파일 세팅
logger = log.setLogging("queryMake")
logger.debug("config file read")

def unzipOracleClient():
    #압축이 풀릴 디렉토리명
    logger.info("instantclient_19_12/ create!")

    with ZipFile('./instantclient-basic-windows.x64-19.12.0.0.0dbru.zip', 'r') as zip:
        zip.extractall()
        logger.info('oracleclient.zip file is unzipped in "instantclient_19_12/" folder')

def queryMake(_textEntry=None):
    logger.info("query make start")

    t_before = threading.Thread(target=beforeExcelToQuery)
    t_before.start()

    t_after = threading.Thread(target=afterExcelToQuery)
    t_after.start()
    t_after.join()
    time = str(datetime.datetime.now())[0:-7]
    _textEntry.insert(END, "[{}] {}".format(time, '쿼리만들기가 완료되었습니다.\n'))

    _dbInsertFlag = ConfigUtil.config.get('oracle', 'database_insert', fallback='N')
    if _dbInsertFlag == 'Y':
        time = str(datetime.datetime.now())[0:-7]
        _textEntry.insert(END, "[{}] {}".format(time, '오라클DB에 insert를 시작합니다.\n'))
        #DBinsert하도록 되어있으면
        t_oracleInsert = threading.Thread(target=oracleInsert, args=(_textEntry,))
        t_oracleInsert.start()
    logger.info("query make end")

def beforeExcelToQuery():

    #DataFrame을 가져온다.
    beforeDf = pd.read_excel('./excel_result/excel_before.xlsx')

    logger.info("before_dataFrame is ready")

    #config에서 테이블명 불러오기
    beforeTableName = ConfigUtil.config.get('oracle', 'expired_table', fallback='END_CONTRACT')

    #쓰기모드로 open
    file = open("./sql/before.sql", 'w', encoding='utf8')
    for index, row in beforeDf.iterrows():#beforeDf의 row를 순회
        # N번쨰 row 가져오기
        # 각 컬럼이 16개나 되기때문에 index로 접근하기위해 .loc사용
        col0 = str(beforeDf.loc[index][0])#연번
        col4 = str(beforeDf.loc[index][4]).strip()#피보험자
        col5 = str(beforeDf.loc[index][5]).strip()#피보험 주민번호
        col10 = str(beforeDf.loc[index][10]).strip().split(" ")[0]#계약체결일
        col11 = str(beforeDf.loc[index][11]).strip().split(" ")[0]#계약소멸일
        col13 = str(beforeDf.loc[index][13]).strip()#모집인명
        col14 = str(beforeDf.loc[index][14]).strip()#모집인주민번호

        # insert 쿼리 생성
        a = f'''INSERT INTO {beforeTableName}( "number_old", EXP_DATE, INS_NM, INS_BIRTH, PLNR_NM, PLNR_BIRTH)VALUES('{col0}', '{col11}','{col4}','{col5}','{col13}','{col14}');
        '''

        # insert 쿼리 파일에 write
        file.write(a.strip() + "\n")
        logger.debug("before:: no." + str(index+1) + " row file write")

    logger.info("before_excel to sql success")
    file.close()

    # TODO 현재까지 파악된 바에 따르면
    # t_before.join()으로 t_before가 끝날때까지 대기하는데
    # _textEntry.insert에서 Thread를 점유하여 처리해야 하므로 대기하게된다.
    # 그래서 t_before가 끝나지 않고 deadlock이 걸린다.
    #time = str(datetime.datetime.now())[0:-7]
    #_textEntry.insert(END, "[{}] {}".format(time, 'before 쿼리만들기가 완료되었습니다.\n'))

def afterExcelToQuery():

    #afterDataFrame을 가져온다.
    afterDf = pd.read_excel('./excel_result/excel_after.xlsx')

    logger.info("after_dataFrame is ready")

    # config 에서 테이블명 불러오기
    afterTableName = ConfigUtil.config.get('oracle', 'concluded_table', fallback='NEW_CONTRACT')

    # 쓰기모드로 open
    file = open("./sql/after.sql", 'w', encoding='utf8')
    for index, row in afterDf.iterrows():#afterDf의 row를 순회
        # N번쨰 row 가져오기
        # 각 컬럼이 16개나 되기때문에 index로 접근하기위해 .loc사용
        col0 = str(afterDf.loc[index][0])  # 연번
        col4 = str(afterDf.loc[index][4]).strip()  # 피보험자
        col5 = str(afterDf.loc[index][5]).strip()  # 피보험 주민번호
        col10 = str(afterDf.loc[index][10]).strip().split(" ")[0]  # 계약체결일
        col11 = str(afterDf.loc[index][11]).strip().split(" ")[0]  # 계약소멸일
        col13 = str(afterDf.loc[index][13]).strip()  # 모집인명
        col14 = str(afterDf.loc[index][14]).strip()  # 모집인주민번호

        #insert 쿼리 생성
        a = f'''INSERT INTO {afterTableName}("number_new", CON_DATE, INS_NM, INS_BIRTH, PLNR_NM, PLNR_BIRTH)VALUES('{col0}','{col10}','{col4}','{col5}','{col13}','{col14}');
        '''

        #insert쿼리 파일에 write
        file.write(a.strip() + "\n")
        logger.debug("after:: no." + str(index+1) + " row file write")

    logger.info("after_excel to sql success")
    file.close()

def oracleInsert(_textEntry=None):
    logger.info("oracle db insert is start")

    # instantclient_19_12/가 없다면
    if not os.path.isdir("./instantclient_19_12"):
        logger.info("instantclient_19_12/ is not exist")
        unzipOracleClient()
        time = str(datetime.datetime.now())[0:-7]
        _textEntry.insert(END, "[{}] {}".format(time, '오라클 접속모듈 압축풀기\n'))
    else:
        # instantclient_19_12/가 비어있다면
        if len(os.listdir("./instantclient_19_12")) == 0:
            logger.info("instantclient_19_12/ is empty")
            unzipOracleClient()
            time = str(datetime.datetime.now())[0:-7]
            _textEntry.insert(END, "[{}] {}".format(time, '오라클 접속모듈 압축풀기\n'))

    # 오라클 클라이언트 디렉토리 (상대경로 -> 절대경로)
    oracleClientDir = './instantclient_19_12'
    oracleClientDir_abs = os.path.abspath(oracleClientDir)

    # 환경변수 설정(oracle 접속을위해 필요
    LOCATION = oracleClientDir_abs
    os.environ["PATH"] = LOCATION + ";" + os.environ["PATH"]
    logger.debug("env setting complete")

    #설정파일 변수에 저장
    _ip = ConfigUtil.config.get('oracle', 'ip', fallback="10.1.1.1")
    _port = int(ConfigUtil.config.get('oracle', 'port', fallback="1521"))
    _service = ConfigUtil.config.get('oracle', 'service', fallback='service')
    _userName = ConfigUtil.config.get('oracle', 'user_name', fallback='admin')
    _userPwd = ConfigUtil.config.get('oracle', 'user_password', fallback='admin')

    #TODO 보안이슈있음 사용자와 비번은 커밋하지말것!

    # _ip = "211.238.147.202"
    # _port = 1521
    # _service = "netis"
    # _userName = ""
    # _userPwd = ""

    #오라클 접속
    dsn = cx_Oracle.makedsn(_ip, _port, _service)
    conn = cx_Oracle.connect(_userName, _userPwd, dsn)
    logger.debug("database connect!")

    cursor = conn.cursor()
    #기존 데이터 삭제
    cursor.execute("DELETE FROM END_CONTRACT")
    cursor.execute("DELETE FROM NEW_CONTRACT")
    conn.commit()

    time = str(datetime.datetime.now())[0:-7]
    _textEntry.insert(END, "[{}] {}".format(time, '기존 테이블 데이터삭제\n'))

    #sql파일을 line 단위로 읽어서 실행한다.
    file1_lines = open("./sql/before.sql", encoding='utf8').readlines()
    for line in file1_lines:
        if line.strip():
            cursor.execute(line.strip().replace(";", ""))
    conn.commit()

    file2_lines = open("./sql/after.sql", encoding='utf8').readlines()
    for line in file2_lines:
        if line.strip():
            cursor.execute(line.strip().replace(";", ""))
    conn.commit()

    conn.close()
    logger.info("oracle db insert is done")

    time = str(datetime.datetime.now())[0:-7]
    _textEntry.insert(END, "[{}] {}".format(time, '데이터 INSERT 성공!\n'))

#queryMake()

#oracleInsert()