import pandas as pd
import numpy as np
import datetime
#import Levenshtein


def dataSheetMerge(df_Sheet1):
    selectedColumnName1 = df_Sheet1.columns

    df_MergedSheet1 = pd.DataFrame(columns=selectedColumnName1)
    df_Sheet1['col2'] = df_Sheet1[selectedColumnName1[1]].apply(lambda x: x[5:7])
    while (df_Sheet1.shape[0] > 0):
        keyValue = [df_Sheet1.iloc[0, 0], df_Sheet1.iloc[0, 1], df_Sheet1.iloc[0, 2]]
        tmpDF = df_Sheet1[df_Sheet1[selectedColumnName1[0]] == keyValue[0]]
        #tmpDF['col2'] = tmpDF[selectedColumnName1[1]].map(lambda x: x[5:7])
        #date1 = df_Sheet1[selectedColumnName1[1]]
        date2 = keyValue[1]
        tmpDF = tmpDF[df_Sheet1['col2'] == date2[5:7]]
        tmpDF = tmpDF[df_Sheet1[selectedColumnName1[2]] == keyValue[2]]
        sumValue = np.sum([float(x) for x in tmpDF[selectedColumnName1[3]].tolist()])
        row = pd.DataFrame([[df_Sheet1.iloc[0, 0], df_Sheet1.iloc[0, 1], df_Sheet1.iloc[0, 2], int(sumValue)]],
                           columns=selectedColumnName1)
        df_MergedSheet1 = pd.concat([df_MergedSheet1, row], ignore_index=True)
        indexList = [x for x in range(tmpDF.shape[0])]
        indexList1 = tmpDF.index.values
        #df_Sheet1 = df_Sheet1.drop(df_Sheet1.index[indexList])
        df_Sheet1 = df_Sheet1.drop(indexList1)
    df_MergedSheet1.index = [int(x) for x in range(df_MergedSheet1.shape[0])]
    return df_MergedSheet1

def dataSheetMatch(dSheet1, dSheet2):

    #print(dSheet1)
    #print(dSheet2)
    matchedDataSheet = pd.merge(dSheet1, dSheet2, how='inner')
    #print(matchedDataSheet)

    dSheet11 = dSheet1.append(matchedDataSheet)
    missedSheet1 = dSheet11.drop_duplicates(keep=False)

    dSheet22 = dSheet2.append(matchedDataSheet)
    missedSheet2 = dSheet22.drop_duplicates(keep=False)
    print([matchedDataSheet.shape, missedSheet1.shape, missedSheet2.shape])


    columnName = missedSheet2.columns
    selectedIndex1 = pd.DataFrame()
    selectedIndex2 = pd.DataFrame()
    for index, row in missedSheet1.iterrows():
        keyValue = row.tolist()
        #print(keyValue)
        tmpDF = missedSheet2.loc[missedSheet2[columnName[0]] == keyValue[0]]
        tmpDF = tmpDF.loc[missedSheet2[columnName[3]] == keyValue[3]]

        for index2, row2 in tmpDF.iterrows():
            targetedValue = row2.tolist()

            datetime1 = datetime.datetime.strptime(keyValue[1], "%Y-%m-%d")
            datetime2 = datetime.datetime.strptime(targetedValue[1], '%Y-%m-%d')

            typeNumber1 = keyValue[2]
            typeNumber2 = targetedValue[2]
            if (abs((datetime1-datetime2).days) < 60 and ((typeNumber1 == typeNumber2) or (typeNumber1.find(typeNumber2) != -1) or (typeNumber2.find(typeNumber1) != -1))):

                selectedIndex1 = selectedIndex1.append(row)
                selectedIndex2 = selectedIndex2.append(row2)

                break

    print([matchedDataSheet.shape, missedSheet1.shape, missedSheet2.shape])
    print([selectedIndex1.shape, selectedIndex2.shape])

    matchedDataSheet = pd.concat([matchedDataSheet, selectedIndex1], ignore_index=True)
    missedSheet1 = pd.concat([missedSheet1, selectedIndex1], ignore_index=True)
    missedSheet2 = pd.concat([missedSheet2, selectedIndex2], ignore_index=True)

    missedSheet1 = missedSheet1.drop_duplicates(keep=False)
    missedSheet2 = missedSheet2.drop_duplicates(keep=False)

    #print([matchedDataSheet.shape, missedSheet1.shape, missedSheet2.shape])
    return matchedDataSheet, missedSheet1, missedSheet2




if __name__ == '__main__':

    strFileFolder = 'C:/Users/GE3154/Desktop/test1/'
    datasheet1 = '贝尔令.xlsx'
    selectedColumnName1 = ['客户名称', '开票日期', '规格', '开票数量']
    datasheet2 = '贝尔令金税.xlsx'
    selectedColumnName2 = ['购方企业名称', '开票日期', '规格', '数量']

    print("Start Load and reshape the datasheet1")
    #Load and reshape the datasheet1
    df_Sheet1 = pd.read_excel(strFileFolder + datasheet1, sheet_name='数据表')
    df_Sheet1 = pd.DataFrame(np.transpose([df_Sheet1[name].tolist() for name in selectedColumnName1]),
                             columns=selectedColumnName1)

    df_Sheet1[selectedColumnName1[0]] = [x.replace('）', ')').replace('（','(') for x in df_Sheet1[selectedColumnName1[0]]]
    df_Sheet1 = df_Sheet1.sort_values(by=selectedColumnName1)
    df_Sheet1 = dataSheetMerge(df_Sheet1)

    print("Start Load and reshape the datasheet2")

    # Load and reshape the datasheet2
    df_Sheet2 = pd.read_excel(strFileFolder + datasheet2, sheet_name='Sheet3')
    df_Sheet2 = pd.DataFrame(np.transpose([df_Sheet2[name].tolist() for name in selectedColumnName2]),
                             columns=selectedColumnName2)
    df_Sheet2[selectedColumnName2[0]] = [x.replace('）', ')').replace('（', '(') for x in
                                         df_Sheet2[selectedColumnName2[0]]]
    df_Sheet2 = df_Sheet2.sort_values(by=selectedColumnName2)
    df_Sheet2 = dataSheetMerge(df_Sheet2)
    df_Sheet2.columns = selectedColumnName1

    outputFile = 'outputMergeItems.xlsx'

    excelWriter = pd.ExcelWriter(outputFile, engine='openpyxl')
    df_Sheet1.to_excel(excel_writer=excelWriter, sheet_name=datasheet1)
    df_Sheet2.to_excel(excel_writer=excelWriter, sheet_name=datasheet2)
    excelWriter.save()
    excelWriter.close()

    print("Merge the two sheet and generate the matched, missed parts")
    # Merge the two sheet and generate the matched, missed parts
    matchedDataSheet, missedSheet1, missedSheet2 = dataSheetMatch(df_Sheet1, df_Sheet2)
    outputFile = 'output.xlsx'

    excelWriter = pd.ExcelWriter(outputFile, engine='openpyxl')
    matchedDataSheet.to_excel(excel_writer=excelWriter, sheet_name='匹配项')
    missedSheet1.to_excel(excel_writer=excelWriter, sheet_name=datasheet1)
    missedSheet2.to_excel(excel_writer=excelWriter, sheet_name=datasheet2)
    excelWriter.save()
    excelWriter.close()
    print("End writing excel")




