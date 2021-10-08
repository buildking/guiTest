from tkinter import *
import combine
import compare
import queryMake
import log
import datetime

if __name__ == '__main__':
    #logger 설정
    logger = log.setLogging("main")
    logger.info("Start!")

    root = Tk()

    root.title("업무 자동화 시스템")
    root.geometry("500x300")#로그표시땜에 늘림
    root.resizable(True, False)#로그표시가 잘안되서 넒이는 늘릴수있게 함

    excelMergeBtn = Button(root, text="엑셀 합치기", command=lambda: combine.excelCombine(resultText))
    excelMergeBtn.pack(padx=2, pady=2, fill='x')

    queryMakeBtn = Button(root, text="쿼리만들기", command=lambda: queryMake.queryMake(resultText))
    queryMakeBtn.pack(padx=2, pady=2, fill='x')

    excelCompareBtn = Button(root, text="엑셀비교하기", command=lambda: compare.excelCompare(resultText))
    excelCompareBtn.pack(padx=2, pady=2, fill='x')

    resultText = Text(root, bg='light gray')
    resultText.pack(padx=2, pady=2, fill='both')

    time = str(datetime.datetime.now())[0:-7]
    resultText.insert(END, "[{}] {}".format(time, '프로그램이 시작되었습니다.\n'))

    root.mainloop()