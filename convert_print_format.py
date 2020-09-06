from openpyxl import load_workbook


def formating():
    path = 'C:/Users/thiag/Desktop/Atlas/test.xlsx'

    loading = load_workbook(filename=path)
    la = loading.active
    la.delete_cols(5)
    la.delete_cols(6)
    la.delete_cols(7)
    la.delete_cols(7)
    la.delete_cols(8, 11)

    la.move_range('E1:E100', rows=0, cols=3)
    la.move_range('G1:G100', rows=0, cols=-2)
    la.move_range('H1:H100', rows=0, cols=-1)

    la['H1'] = 'Tipo de NF'
    la['I1'] = 'Informar NF gerada'

    for i in range(2, 100, 1):
        if (i % 2) != 0:
            measure_code_text = 'E'+str(i)
            la.move_range(measure_code_text, rows=-1, cols=2)
            measure_text = 'G'+str(i)
            la.move_range(measure_text, rows=-1, cols=1)
        else:
            measure_text = 'G'+str(i)
            la.move_range(measure_text, rows=0, cols=-1)

    for j in range(3, 100, 1):
        la.delete_rows(j)

    loading.save(path)


formating()
