import pandas as pd

#before, after excel file compare

def excelCompare():

    # dataframe declare
    beforeDataFrame = pd.DataFrame()
    afterDataFrame = pd.DataFrame()

    beforeCore = pd.DataFrame()
    afterCore = pd.DataFrame()

    finalResultBefore = pd.DataFrame()
    finalResultAfter = pd.DataFrame()

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

    print('*'*50, coreRowNum, '*'*50)

    lastNum = []
    lastNum.clear()

    #조건을 통과한 값의 b'계약소멸일'-a'계약체결일'이 180이내인지 추려냄
    for i, j in coreRowNum:
        #날짜 차이가 180일 이내인지 비교
        if abs(int(str(pd.to_datetime(str(afterDataFrame.loc[j,'계약체결일']))-pd.to_datetime(str(beforeDataFrame.loc[i,'계약소멸일']))).split()[0])) < 181 :
            #aSubB = abs(a[계약체결일] - b[계약소멸일])
            aSubB = int(str(pd.to_datetime(str(afterDataFrame.loc[j, '계약체결일'])) - pd.to_datetime(
                str(beforeDataFrame.loc[i, '계약소멸일']))).split()[0])
            lastNum.append([i,j, aSubB])

    lastExcel = pd.DataFrame(lastNum, columns=['beforeExcel','afterExcel', '차이'])
    print(lastNum)
    lastExcel.to_excel('./excel_final/excel_final_rows.xlsx', index=False)

    #최종적으로 걸러진 row값을 이용해 excel_final 생성
    finalExcel = pd.DataFrame(columns=[['소멸계약<이동 전>', '소멸계약<이동 전>', '소멸계약<이동 전>', '소멸계약<이동 전>', '소멸계약<이동 전>', '소멸계약<이동 전>', '소멸계약<이동 전>', '신규계약<이동 후>', '신규계약<이동 후>', '신규계약<이동 후>', '신규계약<이동 후>', '신규계약<이동 후>', '신규계약<이동 후>', '신규계약<이동 후>', '신규계약<이동 후>', '신규계약<이동 후>', '차이'],['회사명', '피보험자', '증권번호', '상품명', '상품분류', '계약소멸일', '상태', '회사명', '피보험자', '증권번호', '상품명', '상품분류', '계약체결일', '상태', '모집인명', '생년월일', '차이']])

    x = 0
    for b, a, diff in lastNum:
        x += 1
        finalExcel.loc[x] = (beforeDataFrame.loc[b, '회사명'], beforeDataFrame.loc[b, '피보험자'], beforeDataFrame.loc[b, '증권번호'], beforeDataFrame.loc[b, '상품명'], beforeDataFrame.loc[b, '상품분류'], beforeDataFrame.loc[b, '계약소멸일'], beforeDataFrame.loc[b, '계약상태'], afterDataFrame.loc[a, '회사명'], afterDataFrame.loc[a, '피보험자'], afterDataFrame.loc[a, '증권번호'], afterDataFrame.loc[a, '상품명'], afterDataFrame.loc[a, '상품분류'], afterDataFrame.loc[a, '계약체결일'], afterDataFrame.loc[a, '계약상태'], afterDataFrame.loc[a, '모집인명'], afterDataFrame.loc[a, '피보험자\n주민등록번호'], diff)

    finalExcel.to_excel('./excel_final/excel_final.xlsx')

