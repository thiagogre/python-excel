import pandas as pd
from openpyxl import load_workbook

excel_file = 'C:/Users/thiag/Desktop/Atlas/QMtest.xlsx'

loading = load_workbook(filename=excel_file)
la = loading.active
la.delete_cols(7)
la.delete_cols(7)
la.delete_cols(7)
la.delete_cols(7)
la.delete_cols(8)
la.delete_cols(8)
la.delete_cols(8)
la.delete_cols(8)
la.delete_cols(5)
loading.save(excel_file)

df1 = pd.read_excel(excel_file, 'Sheet1')
df2 = pd.DataFrame(df1)
empDfObj = pd.DataFrame(df1)
modDf = empDfObj.dropna()
byNoteDf = (modDf.set_index(['Nota', 'Material', 'Iníc.planj', 'Texto breve material','Texto das medidas','Texto de code medida']))
byNoteDf.to_excel(excel_file)