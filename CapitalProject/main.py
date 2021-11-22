import tkinter
from tkinter import *
import combine
import log
import datetime
import sqlite3
import DbUtil

DbUtil.DbUtil.createTable(None)

def process(textEntry=None):
    #1. 엑셀 합치기
    combine.excelCombine(resultText)
    #2. sqlite insert
    #3. 엑셀 비교
    
if __name__ == '__main__':
    #logger 설정
    logger = log.setLogging("main")
    logger.info("Start!")

    root = Tk()

    root.title("업무 자동화 시스템")
    root.geometry("500x300")
    root.resizable(True, False)

    buttonBox = tkinter.Frame(root)
    buttonBox.pack(side='top', fill='x')
    textBox = tkinter.Frame(root)
    textBox.pack(side='bottom', fill='both')

    jobStartBtn = Button(buttonBox, text="작업시작", command=lambda: process(resultText), height=3, bg='light blue')
    jobStartBtn.pack(padx=2, pady=2, fill='x')

    #출력화면에 스크롤바 추가
    scrollbar = tkinter.Scrollbar(textBox)
    scrollbar.pack(side='right', fill='y')

    resultText = Text(textBox, bg='light gray', yscrollcommand=scrollbar.set)
    resultText.pack(padx=2, pady=2, fill='both')

    scrollbar["command"] = resultText.yview

    time = str(datetime.datetime.now())[0:-7]
    resultText.insert(END, f"[{time}] 프로그램이 시작되었습니다.\n")
    resultText.see(END)

    root.mainloop()