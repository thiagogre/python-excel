import pandas as pd
import collections


# Organizing data
def findEvenText(makeDictRecurrence):
    dictionary = {}
    for index in makeDictRecurrence:
        dictionary[index] = collections.Counter(makeDictRecurrence)[index]
    orderedList = sorted(dictionary.items(), key=lambda x: x[1], reverse=True)
    return orderedList


# Exporting on EXCEL
def writeExcel(listaToExcel):
    df_to_excel = pd.DataFrame(listaToExcel)
    path_to_excel = f'C:/Users/thiag/Desktop/Atlas/geral/{listaToExcel[0]}.xlsx'
    df_to_excel.to_excel(path_to_excel, header=False, startcol=0, index=None)
