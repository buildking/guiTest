from tkinter import *
from os import *
import tkinter.messagebox as msbox
import pandas as pd

root = Tk()

#####################################파일 읽어오기##################################

#before read
beforeFileList = listdir('excel_before')

print(beforeFileList)

readExcel = []

for tmp in beforeFileList:

    tmpExcel = pd.read_excel(f'./excel_before/{tmp}')

    readExcel.append(tmpExcel)

#print(readExcel)

#after read
beforeFileList2 = listdir('excel_after')

#print(beforeFileList2)

readExcel2 = []

for tmp2 in beforeFileList2:

    tmpExcel2 = pd.read_excel(f'./excel_after/{tmp2}')

    readExcel2.append(tmpExcel2)

#print(readExcel2)

####################################################################################################################

#에러메시지 출력(프로세스는 조금 더 고민해야될듯)
def combineError():
    pass

#엑셀 합치기
def excelCombine():
    global readExcel

    #before merge
    if len(readExcel) == 0 or len(readExcel) == 1:
        combineError()
        return
    minusexcel = len(readExcel) - 1
    for tmp in range(minusexcel):
        combineExcel = pd.merge(readExcel[tmp], readExcel[tmp+1], how='outer')
    combineExcel.to_excel('./excel_result/excel_before.xlsx', index=False)

    #after merge
    if len(readExcel2) == 0 or len(readExcel2) == 1:
        combineError()
        return
    minusexcel2 = len(readExcel2) - 1
    for tmp in range(minusexcel2):
        combineExcel2 = pd.merge(readExcel2[tmp], readExcel2[tmp+1], how='outer')
    combineExcel2.to_excel('./excel_result/excel_after.xlsx', index=False)


#before, after excel file compare

#dataframe declare
beforeDataFrame = pd.DataFrame()
afterDataFrame = pd.DataFrame()

beforeCore = pd.DataFrame()
afterCore = pd.DataFrame()

finalResultBefore = pd.DataFrame()
finalResultAfter = pd.DataFrame()

def excelCompare():
    global beforeDataFrame, afterDataFrame, lastNum

    beforeDataFrame = pd.read_excel('./excel_result/excel_before.xlsx')
    afterDataFrame = pd.read_excel('./excel_result/excel_after.xlsx')

    # print(beforeDataFrame)
    # print(afterDataFrame)

    #비교 통과한 before data와 after data의 row 저장(append로)
    coreRowNum = []
    coreRowNum.clear()


    # for pdb, rowBefore in beforeDataFrame.iterrows():
    #     print('되냐??'*10)
    #     print(pdb,'뎃데로게'*10, rowBefore)
    #     print('야지막으로, ', rowBefore['피보험자\n주민등록번호'])

    for pdb, rowBefore in beforeDataFrame.iterrows():
        print(pdb, rowBefore)
        for pda, rowAfter in afterDataFrame.iterrows():
            print(pda, rowAfter)
            if str(rowBefore['피보험자']) == str(rowAfter['피보험자']):
                if str(rowBefore['모집인명']) == str(rowAfter['모집인명']):
                    if str(rowBefore['피보험자\n주민등록번호'])[:5] == str(rowAfter['피보험자\n주민등록번호'])[:5]:
                        if str(rowBefore['모집인 주민번호'])[:5] == str(rowAfter['모집인 주민번호'])[:5]:
                            coreRowNum.append([pdb, pda])
###################################################################################
#

    print('*'*50, coreRowNum, '*'*50)

    lastNum = []
    lastNum.clear()

    for i, j in coreRowNum:
        #날짜 차이가 180일 이내인지 비교
        if int(str(pd.to_datetime(str(afterDataFrame.loc[j,'계약체결일']))-pd.to_datetime(str(beforeDataFrame.loc[i,'계약체결일']))).split()[0]) < 181 :
            lastNum.append([i,j])

    lastExcel = pd.DataFrame(lastNum, columns=['beforeExcel','afterExcel'])
    print(lastNum)
    lastExcel.to_excel('./excel_final/excel_final.xlsx', index=False)




#query making
# def queryMake():






root.title("업무 자동화 시스템")
root.geometry("400x300")
root.resizable(False, False)

excelMergeBtn = Button(root, text="엑셀 합치기", command= lambda : excelCombine())
excelMergeBtn.pack(padx=2, pady=2, fill='x')

queryMakeBtn = Button(root, text="쿼리만들기", command=None)
queryMakeBtn.pack(padx=2, pady=2, fill='x')

excelCompareBtn = Button(root, text="엑셀비교하기", command= lambda : excelCompare())
excelCompareBtn.pack(padx=2, pady=2, fill='x')

resultText = Text(root, bg='light gray')
resultText.pack(padx=2, pady=2, fill='both')

root.mainloop()
