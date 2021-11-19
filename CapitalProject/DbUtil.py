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
                _param["numberOld"], _param["expDate"], _param["insNm"], _param["insBirth"], _param["plnrNm"],
                _param["plnrBirth"]))
        self.conn.commit()

    def insertEndContract(self, _paramList):
        cursor = self.getCursor()
        for _param in _paramList:
            cursor.execute("INSERT INTO END_CONTRACT VALUES (?, ?, ?, ?, ?, ?)", (
            _param["numberOld"], _param["conDate"], _param["insNm"], _param["insBirth"], _param["plnrNm"],
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
        pass
