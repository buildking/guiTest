import pandas as pd
import config
import cx_Oracle
import log

#로그파일 세팅
logger = log.setLogging("queryMake")
logger.debug("config file read")

#config에서 oracle 섹션을 읽는다.
_config = config.ConfigUtil()
_oracleSetting = _config.getOracleSetting()

def queryMake():
    beforeExcelToQuery()
    logger.info("")
    logger.info("--->before is Done<---")
    logger.info("")
    afterExcelToQuery()
    logger.info("")
    logger.info("--->after is Done<---")
    logger.info("")

def beforeExcelToQuery():

    #DataFrame을 가져온다.
    beforeDf = pd.read_excel('./excel_result/excel_before.xlsx')

    logger.info("before_dataFrame is ready")

    #config에서 테이블명 불러오기
    beforeTableName = _oracleSetting['expired_table']

    #쓰기모드로 open
    file = open("./sql/before.sql", 'w', encoding='utf8')
    for index, row in beforeDf.iterrows():#beforeDf의 row를 순회
        logger.debug("no." + str(index) + " row read!")
        # N번쨰 row 가져오기
        # 각 컬럼이 16개나 되기때문에 index로 접근하기위해 .loc사용
        col1 = beforeDf.loc[index][0]
        col2 = beforeDf.loc[index][1]
        col3 = beforeDf.loc[index][2]
        col4 = beforeDf.loc[index][3]
        col5 = beforeDf.loc[index][4]
        col6 = beforeDf.loc[index][5]
        col7 = beforeDf.loc[index][6]
        col8 = beforeDf.loc[index][7]
        col9 = beforeDf.loc[index][8]
        col10 = beforeDf.loc[index][9]
        col11 = beforeDf.loc[index][10]
        col12 = beforeDf.loc[index][11]
        col13 = beforeDf.loc[index][12]
        col14 = beforeDf.loc[index][13]
        col15 = beforeDf.loc[index][14]
        col16 = beforeDf.loc[index][15]

        # insert 쿼리 생성
        a = f'''INSERT INTO {beforeTableName}( col1, col2, col3, col4, col5, col6, col7, col8, col9, col10, col11, col12, col13, col14, col15, col16)
            VALUE('{col1}', '{col2}','{col3}','{col4}','{col5}','{col6}','{col7}','{col8}','{col9}','{col10}','{col11}','{col12}','{col13}','{col14}','{col15}','{col16}');
        '''

        # insert 쿼리 파일에 write
        file.write(a)
        file.write("\n")
        logger.info("before:: no." + str(index+1) + " row file write")

    logger.info("before_excel to sql success")
    file.close()

def afterExcelToQuery():

    #afterDataFrame을 가져온다.
    afterDf = pd.read_excel('./excel_result/excel_after.xlsx')

    logger.info("after_dataFrame is ready")

    # config 에서 테이블명 불러오기
    afterTableName = _oracleSetting['concluded_table']

    # 쓰기모드로 open
    file = open("./sql/after.sql", 'w', encoding='utf8')
    for index, row in afterDf.iterrows():#afterDf의 row를 순회
        logger.debug("no." + str(index) + " row read!")
        # N번쨰 row 가져오기
        # 각 컬럼이 16개나 되기때문에 index로 접근하기위해 .loc사용
        col1 = afterDf.loc[index][0]
        col2 = afterDf.loc[index][1]
        col3 = afterDf.loc[index][2]
        col4 = afterDf.loc[index][3]
        col5 = afterDf.loc[index][4]
        col6 = afterDf.loc[index][5]
        col7 = afterDf.loc[index][6]
        col8 = afterDf.loc[index][7]
        col9 = afterDf.loc[index][8]
        col10 = afterDf.loc[index][9]
        col11 = afterDf.loc[index][10]
        col12 = afterDf.loc[index][11]
        col13 = afterDf.loc[index][12]
        col14 = afterDf.loc[index][13]
        col15 = afterDf.loc[index][14]
        col16 = afterDf.loc[index][15]

        #insert 쿼리 생성
        a = f'''INSERT INTO {afterTableName}( col1, col2, col3, col4, col5, col6, col7, col8, col9, col10, col11, col12, col13, col14, col15, col16)
            VALUE('{col1}', '{col2}','{col3}','{col4}','{col5}','{col6}','{col7}','{col8}','{col9}','{col10}','{col11}','{col12}','{col13}','{col14}','{col15}','{col16}');
        '''

        #insert쿼리 파일에 write
        file.write(a)
        file.write("\n")
        logger.info("after:: no." + str(index+1) + " row file write")

    logger.info("after_excel to sql success")
    file.close()


def oracleConnect():
    dsn = cx_Oracle.makedsn("IP", "PORT:int", "SID")
    conn = cx_Oracle.connect("USER_ID", "PASSWORD", dsn)

    cursor = conn.cursor()
    cursor.execute("select 1 from dual")

    for data in cursor:
        print(data)

    conn.close()

queryMake()