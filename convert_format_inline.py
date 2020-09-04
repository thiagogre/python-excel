import pandas as pd
from openpyxl import load_workbook
from collections import Counter
import numpy as np

excel_file = 'C:/Users/thiag/Desktop/Atlas/QMtest.xlsx'
excel_file_inline = 'C:/Users/thiag/Desktop/Atlas/inline.xlsx'
loading = load_workbook(filename=excel_file_inline)
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
loading.save(excel_file_inline)


dfInline = pd.read_excel(excel_file_inline)
loading1 = load_workbook(filename=excel_file_inline)
la1 = loading1.active
j = 2
repeated_notes = {}
copy = ''
aux = 0
for index in dfInline['Nota']:

    repeated_notes[index] = Counter(dfInline['Nota'])[index]

    if (repeated_notes[index] == 2) and (copy != index):
        textoDasMedidas = 'E'+str(j)
        la1.move_range(textoDasMedidas, rows=1, cols=2)
        textoCodeMedidas = 'F'+str(j)
        la1.move_range(textoCodeMedidas, rows=1, cols=2)
        copy = index

    if (repeated_notes[index] == 3) and (copy != index):
        if aux == 0:
            textoDasMedidas = 'E'+str(j+1)
            la1.move_range(textoDasMedidas, rows=-1, cols=2)
            textoCodeMedidas = 'F'+str(j+1)
            la1.move_range(textoCodeMedidas, rows=-1, cols=3)
            aux += 1
            copy = index
            if aux == 1:
                textoDasMedidas = 'E'+str(j+2)
                la1.move_range(textoDasMedidas, rows=-2, cols=3)
                textoCodeMedidas = 'F'+str(j+2)
                la1.move_range(textoCodeMedidas, rows=-2, cols=4)
                aux = 0

    if (repeated_notes[index] == 4) and (copy != index):
        if aux == 0:
            textoDasMedidas = 'E'+str(j+1)
            la1.move_range(textoDasMedidas, rows=-1, cols=2)
            textoCodeMedidas = 'F'+str(j+1)
            la1.move_range(textoCodeMedidas, rows=-1, cols=2)
            aux += 1
            copy = index
            if aux == 1:
                textoDasMedidas = 'E'+str(j+2)
                la1.move_range(textoDasMedidas, rows=-2, cols=4)
                textoCodeMedidas = 'F'+str(j+2)
                la1.move_range(textoCodeMedidas, rows=-2, cols=4)
                aux += 1
                if aux == 2:
                    textoDasMedidas = 'E'+str(j+3)
                    la1.move_range(textoDasMedidas, rows=-3, cols=6)
                    textoCodeMedidas = 'F'+str(j+3)
                    la1.move_range(textoCodeMedidas, rows=-3, cols=6)
                    aux = 0

    if (repeated_notes[index] == 5) and (copy != index):
        if aux == 0:
            textoDasMedidas = 'E'+str(j+1)
            la1.move_range(textoDasMedidas, rows=-1, cols=2)
            textoCodeMedidas = 'F'+str(j+1)
            la1.move_range(textoCodeMedidas, rows=-1, cols=2)
            aux += 1
            copy = index
            if aux == 1:
                textoDasMedidas = 'E'+str(j+2)
                la1.move_range(textoDasMedidas, rows=-2, cols=4)
                textoCodeMedidas = 'F'+str(j+2)
                la1.move_range(textoCodeMedidas, rows=-2, cols=4)
                aux += 1
                if aux == 2:
                    textoDasMedidas = 'E'+str(j+3)
                    la1.move_range(textoDasMedidas, rows=-3, cols=6)
                    textoCodeMedidas = 'F'+str(j+3)
                    la1.move_range(textoCodeMedidas, rows=-3, cols=6)
                    aux += 1
                    if aux == 3:
                        textoDasMedidas = 'E'+str(j+4)
                        la1.move_range(textoDasMedidas, rows=-4, cols=8)
                        textoCodeMedidas = 'F'+str(j+4)
                        la1.move_range(textoCodeMedidas, rows=-4, cols=8)
                        aux = 0
    if (repeated_notes[index] == 6) and (copy != index):
        if aux == 0:
            textoDasMedidas = 'E'+str(j+1)
            la1.move_range(textoDasMedidas, rows=-1, cols=2)
            textoCodeMedidas = 'F'+str(j+1)
            la1.move_range(textoCodeMedidas, rows=-1, cols=2)
            aux += 1
            copy = index
            if aux == 1:
                textoDasMedidas = 'E'+str(j+2)
                la1.move_range(textoDasMedidas, rows=-2, cols=4)
                textoCodeMedidas = 'F'+str(j+2)
                la1.move_range(textoCodeMedidas, rows=-2, cols=4)
                aux += 1
                if aux == 2:
                    textoDasMedidas = 'E'+str(j+3)
                    la1.move_range(textoDasMedidas, rows=-3, cols=6)
                    textoCodeMedidas = 'F'+str(j+3)
                    la1.move_range(textoCodeMedidas, rows=-3, cols=6)
                    aux += 1
                    if aux == 3:
                        textoDasMedidas = 'E'+str(j+4)
                        la1.move_range(textoDasMedidas, rows=-4, cols=8)
                        textoCodeMedidas = 'F'+str(j+4)
                        la1.move_range(textoCodeMedidas, rows=-4, cols=8)
                        aux += 1
                        if aux == 4:
                            textoDasMedidas = 'E'+str(j+5)
                            la1.move_range(textoDasMedidas, rows=-5, cols=10)
                            textoCodeMedidas = 'F'+str(j+5)
                            la1.move_range(textoCodeMedidas, rows=-5, cols=10)
                            aux = 0
    if (repeated_notes[index] == 7) and (copy != index):
        if aux == 0:
            textoDasMedidas = 'E'+str(j+1)
            la1.move_range(textoDasMedidas, rows=-1, cols=2)
            textoCodeMedidas = 'F'+str(j+1)
            la1.move_range(textoCodeMedidas, rows=-1, cols=2)
            aux += 1
            copy = index
            if aux == 1:
                textoDasMedidas = 'E'+str(j+2)
                la1.move_range(textoDasMedidas, rows=-2, cols=4)
                textoCodeMedidas = 'F'+str(j+2)
                la1.move_range(textoCodeMedidas, rows=-2, cols=4)
                aux += 1
                if aux == 2:
                    textoDasMedidas = 'E'+str(j+3)
                    la1.move_range(textoDasMedidas, rows=-3, cols=6)
                    textoCodeMedidas = 'F'+str(j+3)
                    la1.move_range(textoCodeMedidas, rows=-3, cols=6)
                    aux += 1
                    if aux == 3:
                        textoDasMedidas = 'E'+str(j+4)
                        la1.move_range(textoDasMedidas, rows=-4, cols=8)
                        textoCodeMedidas = 'F'+str(j+4)
                        la1.move_range(textoCodeMedidas, rows=-4, cols=8)
                        aux += 1
                        if aux == 4:
                            textoDasMedidas = 'E'+str(j+5)
                            la1.move_range(textoDasMedidas, rows=-5, cols=10)
                            textoCodeMedidas = 'F'+str(j+5)
                            la1.move_range(textoCodeMedidas, rows=-5, cols=10)
                            aux += 1
                            if aux == 5:
                                textoDasMedidas = 'E'+str(j+6)
                                la1.move_range(textoDasMedidas, rows=-6, cols=12)
                                textoCodeMedidas = 'F'+str(j+6)
                                la1.move_range(textoCodeMedidas, rows=-6, cols=12)
                                aux = 0
    if (repeated_notes[index] == 8) and (copy != index):
        if aux == 0:
            textoDasMedidas = 'E'+str(j+1)
            la1.move_range(textoDasMedidas, rows=-1, cols=2)
            textoCodeMedidas = 'F'+str(j+1)
            la1.move_range(textoCodeMedidas, rows=-1, cols=2)
            aux += 1
            copy = index
            if aux == 1:
                textoDasMedidas = 'E'+str(j+2)
                la1.move_range(textoDasMedidas, rows=-2, cols=4)
                textoCodeMedidas = 'F'+str(j+2)
                la1.move_range(textoCodeMedidas, rows=-2, cols=4)
                aux += 1
                if aux == 2:
                    textoDasMedidas = 'E'+str(j+3)
                    la1.move_range(textoDasMedidas, rows=-3, cols=6)
                    textoCodeMedidas = 'F'+str(j+3)
                    la1.move_range(textoCodeMedidas, rows=-3, cols=6)
                    aux += 1
                    if aux == 3:
                        textoDasMedidas = 'E'+str(j+4)
                        la1.move_range(textoDasMedidas, rows=-4, cols=8)
                        textoCodeMedidas = 'F'+str(j+4)
                        la1.move_range(textoCodeMedidas, rows=-4, cols=8)
                        aux += 1
                        if aux == 4:
                            textoDasMedidas = 'E'+str(j+5)
                            la1.move_range(textoDasMedidas, rows=-5, cols=10)
                            textoCodeMedidas = 'F'+str(j+5)
                            la1.move_range(textoCodeMedidas, rows=-5, cols=10)
                            aux += 1
                            if aux == 5:
                                textoDasMedidas = 'E'+str(j+6)
                                la1.move_range(textoDasMedidas, rows=-6, cols=12)
                                textoCodeMedidas = 'F'+str(j+6)
                                la1.move_range(textoCodeMedidas, rows=-6, cols=12)
                                aux += 1
                                if aux == 6:
                                    textoDasMedidas = 'E'+str(j+7)
                                    la1.move_range(textoDasMedidas, rows=-7, cols=14)
                                    textoCodeMedidas = 'F'+str(j+7)
                                    la1.move_range(textoCodeMedidas, rows=-7, cols=14)
                                    aux = 0
    j += 1

loading1.save(excel_file_inline)
df = pd.read_excel(excel_file_inline)
df['Texto das medidas'].replace('', np.nan, inplace=True)
df.dropna(subset=['Texto das medidas'], inplace=True)
df.rename(columns={
    'Unnamed: 6': 'Tipo NF',
    'Unnamed: 7': 'Emissão NF',
    'Unnamed: 8': 'Retorno NF',
    'Unnamed: 9': 'Conhecimento',
    'Unnamed: 10': 'Valor',
    'Unnamed: 11': 'Conhecimento',
    'Unnamed: 12': 'Relatório',
    'Unnamed: 13': 'Conhecimento',
    'Unnamed: 14': 'Material',
    'Unnamed: 15': 'Avaliar',
    'Unnamed: 16': '',
    'Unnamed: 17': '',
    'Unnamed: 18': '',
    'Unnamed: 19': '',
}, inplace=True)

df.to_excel(excel_file_inline)