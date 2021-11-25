import sqlite3

class DbUtil():
    def __init__(self, _dbFile):
        self.conn = sqlite3.connect(_dbFile)

    def getCursor(self):
        return self.conn.cursor()

    def insertNewContract(self, _paramList):
        cursor = self.getCursor()
        for _param in _paramList:
            cursor.execute("INSERT INTO NEW_CONTRACT VALUES (?, ?, ?, ?, ?, ?)", (
                _param["numberNew"], _param["newDate"], _param["insNm"], _param["insBirth"], _param["plnrNm"],
                _param["plnrBirth"]))
        self.conn.commit()

    def insertEndContract(self, _paramList):
        cursor = self.getCursor()
        for _param in _paramList:
            cursor.execute("INSERT INTO END_CONTRACT VALUES (?, ?, ?, ?, ?, ?)", (
            _param["numberEnd"], _param["endDate"], _param["insNm"], _param["insBirth"], _param["plnrNm"],
            _param["plnrBirth"]))
        self.conn.commit()

    def selectNewContract(self):
        cursor = self.conn.cursor()
        cursor.execute("SELECT * FROM NEW_CONTRACT")
        rows = cursor.fetchall()
        return rows

    def selectEndContract(self):
        cursor = self.conn.cursor()
        cursor.execute("SELECT * FROM END_CONTRACT")
        rows = cursor.fetchall()
        return rows

    def deleteNewContract(self):
        cursor = self.getCursor()
        cursor.execute("DELETE FROM NEW_CONTRACT")
        self.conn.commit()

    def deleteEndContract(self):
        cursor = self.getCursor()
        cursor.execute("DELETE FROM END_CONTRACT")
        self.conn.commit()

    def selectCompareResult(self):
        cursor = self.getCursor()
        cursor.execute("""
                   SELECT * 
                     FROM END_CONTRACT END, NEW_CONTRACT NEW
                    WHERE TRIM(END.PLNR_NM) = TRIM(NEW.PLNR_NM)
                      AND TRIM(END.PLNR_BIRTH) = TRIM(NEW.PLNR_BIRTH)
                      AND TRIM(END.INS_NM) = TRIM(NEW.INS_NM)
                      AND TRIM(END.INS_BIRTH) = TRIM(NEW.INS_BIRTH)
                      AND ABS(julianday(END.END_DATE) - julianday(NEW.NEW_DATE)) < 180""")
        rows = cursor.fetchall()
        return rows

    def createTable(self):
        #KniaData = sqlite3.connect('./db/knia.db', isolation_level=None)
        #knDB = KniaData.cursor()
        knDB = self.getCursor()

        readSql = open('./sql/table.sql')

        for _ in range(2):
            knDB.execute(readSql.readline())

        readSql.close()

        #반드시 execute하고선
        self.conn.commit()

        #KniaData.close()

    def databaseClose(self):
        self.conn.close()
