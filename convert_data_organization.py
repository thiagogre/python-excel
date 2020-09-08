import pandas as pd
from utils import findEvenText, writeExcel

excel_file = 'C:/Users/thiag/Desktop/Atlas/DADOS.xlsx'
df1 = pd.read_excel(excel_file, 'QM12')

provider_name = []
material_service = []

nf_service = []
nf_guarantee = []
nfs = []

flawless = []
approved = []
damaged = []
damaged_material = []
broken_seal_material = []
imported_material = []
values_list = []
value_guarantee = []
value_wasted = []
value_service = []

qm = []

service_count = 0
row = 0
filtro_flawless = 0
y = 0

# Importing data from excel
for measure_text_code in df1['Texto de code medida']:
    if measure_text_code == 'Enviar para concerto externo':
        provider_name.append(df1.loc[y, 'Texto das medidas'])
        material_service.append(df1.loc[y, 'Material'])
        service_count += 1
    if measure_text_code == 'Emitir Nota Fiscal':
        if df1.loc[y, 'Texto das medidas'] == 'NF PARA CONSERTO':
            nf_service.append(df1.loc[y, 'Texto das medidas'])
        if df1.loc[y, 'Texto das medidas'] == 'NF EM GARANTIA':
            nf_guarantee.append(df1.loc[y, 'Texto das medidas'])
        row += 1
    y += 1
y = 0

for measure_text in df1['Texto das medidas']:
    df1_material = df1.loc[y, 'Material']
    if measure_text == 'MATERIAL SEM DEFEITO':
        flawless.append(df1_material)
    if measure_text == 'SEM DEFEITO':
        flawless.append(df1_material)
    if measure_text == 'MATERIAL APROVADO':
        approved.append(df1_material)
    if measure_text == 'MATERIAL AVARIADO':
        damaged.append(df1_material)
    if measure_text == 'MATERIAL COM DEFEITO':
        damaged_material.append(df1_material)
    if measure_text == 'MATERIAL COM DEFEITO':
        damaged_material.append(df1_material)
    if measure_text == 'MATERIAL COM LACRE ROMPIDO':
        broken_seal_material.append(df1_material)
    if measure_text == 'MATERIAL IMPORTADO':
        imported_material.append(df1_material)
    try:
        if (measure_text.startswith('VALOR')) is True:
            values_list.append(measure_text)
    except:
        continue
    y += 1
y = 0

provider_in_order = findEvenText(provider_name)
provider_in_order.insert(
    0, ('NOME DO FORNECEDOR', 'QUANTIDADE DE MATERIAIS ENVIADOS'))

send_servive_material = findEvenText(material_service)
send_servive_material.insert(
    0, ('MATERIAL ENVIADO PARA CONSERTO EXTERNO', 'QUANTIDADE'))

nf_service_total = len(nf_service)
nf_guarantee_total = len(nf_guarantee)
nfs.append(
    ['Quantidade de materiais enviados para service externo', service_count])
nfs.append(['Total de nfs geradas', len(nf_service+nf_guarantee)])
nfs.append(['nfs em espera', service_count - len(nf_service+nf_guarantee)])
nfs.append(['Quantidade de nfs de CONSERTO', nf_service_total])
nfs.append(['Quantidade de nfs de GARANTIA', nf_guarantee_total])

flawless_material_total = findEvenText(flawless)
flawless_material_total.insert(0, ('MATERIAL SEM DEFEITO', 'QUANTIDADE'))

approved_material_total = findEvenText(approved)
approved_material_total.insert(0, ('MATERIAL APROVADO', 'QUANTIDADE'))

wasted_material_total = findEvenText(damaged)
wasted_material_total.insert(0, ('MATERIAL AVARIADO', 'QUANTIDADE'))

damaged_material_total = findEvenText(damaged_material)
damaged_material_total.insert(0, ('MATERIAL DEFEITO', 'QUANTIDADE'))

broken_seal_material_total = findEvenText(broken_seal_material)
broken_seal_material_total.insert(
    0, ('MATERIAL COM LACRE ROMPIDO', 'QUANTIDADE'))

imported_material_total = findEvenText(imported_material)
imported_material_total.insert(0, ('MATERIAL IMPORTADO', 'QUANTIDADE'))

value_total = findEvenText(values_list)
value_total.insert(0, ('TODOS VALORES', 'QUANTIDADE'))

for value in values_list:
    if (value.endswith('EM GARANTIA')) is True:
        value_guarantee.append(value)
value_guarantee_total = findEvenText(value_guarantee)
value_guarantee_total.insert(0, ('VALOR EM GARANTIA', 'QUANTIDADE'))

for value in values_list:
    if (value.endswith('SUCATADO')) is True:
        value_wasted.append(value)
value_wasted_total = findEvenText(value_wasted)
value_wasted_total.insert(0, ('VALOR EM SUCATA', 'QUANTIDADE'))

for value in values_list:
    if (value.endswith('DE CONSERTO')) is True:
        value_service.append(value)
value_service_total = findEvenText(value_service)
value_service_total.insert(0, ('VALOR EM CONSERTO', 'QUANTIDADE'))

for note in df1['Nota']:
    qm.append(note)
qm_total = findEvenText(qm)
qm_total.insert(0, ('QM', 'QUANTIDADE'))

writeExcel(send_servive_material)
writeExcel(provider_in_order)
writeExcel(nfs)
writeExcel(flawless_material_total)
writeExcel(approved_material_total)
writeExcel(wasted_material_total)
writeExcel(damaged_material_total)
writeExcel(broken_seal_material_total)
writeExcel(imported_material_total)
writeExcel(value_total)
writeExcel(value_guarantee_total)
writeExcel(value_wasted_total)
writeExcel(value_service_total)
writeExcel(qm_total)
