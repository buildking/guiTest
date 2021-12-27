import tkinter
from tkinter import *
import wemade_log
import datetime
from wemade_excel_compare import *


def process(textEntry=None):
    xlCompare(resultText)


if __name__ == '__main__':
    #logger 설정
    logger = wemade_log.setLogging("main")
    logger.debug("Start!")

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


    def close_window():
        logger.info("program exit")
        root.destroy()
    root.protocol("WM_DELETE_WINDOW", close_window)
    root.mainloop()


    #error massage
    #합칠 파일이 부족할 때 경고문
    def shortError(state):
        messagebox.showwarning("파일 개수 오류", f"{state}폴더에는 엑셀파일을 반드시 2개 넣어야 합니다.")
        logger.debug(f"{state}folder's file must be 2")
        print(f"{state}folder's file must be 2")

    #파일인식 에러
    def inputError(state=None):
        #엑셀파일 인식이 안된다면, 에러메세지 띄우기
        messagebox.showwarning("파일 인식 불가", f"input폴더의 파일을 확인해주세요.")
        logger.debug("excel file don't observed")
        print("excel file don't observed")