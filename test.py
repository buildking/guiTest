import tkinter

import googletrans
import tkinter.ttk as ttk
from tkinter import *


trs = googletrans.Translator()

def multiTranslate(income, lang):
    #번역 결과
    trsbox = trs.translate(income, dest=f"{lang}")
    trsReturn = (str(trsbox).split()[2])[5:-1]
    return trsReturn

transObject = ['korean', 'english', 'french', 'german', 'italian', 'spanish', 'japanese', 'chinese (traditional)', 'chinese (simplified)']

def transResult():
    transSubject = inputTxt.get()

    inputTxt.delete(0, END)
    resultTxt.delete('1.0', END)

    for i in transObject:
        resultTxt.insert(END, f'\n{i}')
        resultTxt.insert(END, f'\n{multiTranslate(transSubject, i)}')
    resultTxt.see(END)

root = Tk()
root.title('auto_translation')
root.geometry('180x320+800+100')

inputFrame = Frame(root)
inputFrame.pack()
resultFrame = Frame(root)
resultFrame.pack()

inputTxt = Entry(inputFrame)
inputTxt.pack(fill='x')
inputTxt.bind("<Return>", transResult)

inputButton = Button(inputFrame, text='번역하기', command=lambda: resultTxt())
inputButton.pack(fill='x')

#scrollbar add
resultScrollbar = tkinter.Scrollbar(resultFrame)
resultScrollbar.pack(side='right', fill='y')

resultTxt = Text(resultFrame, yscrollcommand=resultScrollbar.set)
resultTxt.pack(fill='both')

resultScrollbar['command'] = resultTxt.yview


root.mainloop()