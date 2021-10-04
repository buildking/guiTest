from tkinter import *
from os import *
import tkinter.messagebox as msbox
import pandas as pd
import combine
import compare

root = Tk()


root.title("업무 자동화 시스템")
root.geometry("400x300")
root.resizable(False, False)

excelMergeBtn = Button(root, text="엑셀 합치기", command= lambda : combine.excelCombine())
excelMergeBtn.pack(padx=2, pady=2, fill='x')

queryMakeBtn = Button(root, text="쿼리만들기", command=None)
queryMakeBtn.pack(padx=2, pady=2, fill='x')

excelCompareBtn = Button(root, text="엑셀비교하기", command= lambda : compare.excelCompare())
excelCompareBtn.pack(padx=2, pady=2, fill='x')

resultText = Text(root, bg='light gray')
resultText.pack(padx=2, pady=2, fill='both')

root.mainloop()
