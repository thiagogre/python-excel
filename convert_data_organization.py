import pandas as pd
from utils import findEvenText, writeExcel

excel_file = 'C:/Users/thiag/Desktop/Atlas/DADOS.xlsx'
df1 = pd.read_excel(excel_file, 'QM12')

nome_fornecedor = []
material_conserto = []

nf_conserto = []
nf_garantia = []
NFs = []

sem_defeito = []
material_aprovado = []
material_avariado = []
material_defeito = []
material_lacre_rompido = []
material_importado = []
valor = []
valor_garantia = []
valor_sucatado = []
valor_conserto = []

qm = []

conserto_contador = 0
linha = 0
filtro_sem_defeito = 0
y = 0

# Importing data from excel
for i in df1['Texto de code medida']:
    if i == 'Enviar para concerto externo':
        nome_fornecedor.append(df1.loc[y, 'Texto das medidas'])
        material_conserto.append(df1.loc[y, 'Material'])
        conserto_contador += 1

    if i == 'Emitir Nota Fiscal':
        if df1.loc[y, 'Texto das medidas'] == 'NF PARA CONSERTO':
            nf_conserto.append(df1.loc[y, 'Texto das medidas'])
        if df1.loc[y, 'Texto das medidas'] == 'NF EM GARANTIA':
            nf_garantia.append(df1.loc[y, 'Texto das medidas'])
        linha += 1
    y += 1
y = 0

for i in df1['Texto das medidas']:
    df1_material = df1.loc[y, 'Material']
    if i == 'MATERIAL SEM DEFEITO':
        sem_defeito.append(df1_material)
    if i == 'SEM DEFEITO':
        sem_defeito.append(df1_material)
    if i == 'MATERIAL APROVADO':
        material_aprovado.append(df1_material)
    if i == 'MATERIAL AVARIADO':
        material_avariado.append(df1_material)
    if i == 'MATERIAL COM DEFEITO':
        material_defeito.append(df1_material)
    if i == 'MATERIAL COM DEFEITO':
        material_defeito.append(df1_material)
    if i == 'MATERIAL COM LACRE ROMPIDO':
        material_lacre_rompido.append(df1_material)
    if i == 'MATERIAL IMPORTADO':
        material_importado.append(df1_material)
    try:
        if (i.startswith('VALOR')) is True:
            valor.append(i)
    except:
        continue
    y += 1
y = 0

fornecedorOrdenado = findEvenText(nome_fornecedor)
fornecedorOrdenado.insert(
    0, ('NOME DO FORNECEDOR', 'QUANTIDADE DE MATERIAIS ENVIADOS'))

materiaisEnviadosConserto = findEvenText(material_conserto)
materiaisEnviadosConserto.insert(
    0, ('MATERIAL ENVIADO PARA CONSERTO EXTERNO', 'QUANTIDADE'))

nfConserto = len(nf_conserto)
nfGarantia = len(nf_garantia)
NFs.append(
    ['Quantidade de materiais enviados para conserto externo', conserto_contador])
NFs.append(['Total de NFs geradas', len(nf_conserto+nf_garantia)])
NFs.append(['NFs em espera', conserto_contador - len(nf_conserto+nf_garantia)])
NFs.append(['Quantidade de NFs de CONSERTO', nfConserto])
NFs.append(['Quantidade de NFs de GARANTIA', nfGarantia])

materiaisSemDefeito = findEvenText(sem_defeito)
materiaisSemDefeito.insert(0, ('MATERIAL SEM DEFEITO', 'QUANTIDADE'))

materialAprovado = findEvenText(material_aprovado)
materialAprovado.insert(0, ('MATERIAL APROVADO', 'QUANTIDADE'))

materialAvariado = findEvenText(material_avariado)
materialAvariado.insert(0, ('MATERIAL AVARIADO', 'QUANTIDADE'))

materialDefeito = findEvenText(material_defeito)
materialDefeito.insert(0, ('MATERIAL DEFEITO', 'QUANTIDADE'))

materialLacreRompido = findEvenText(material_lacre_rompido)
materialLacreRompido.insert(0, ('MATERIAL COM LACRE ROMPIDO', 'QUANTIDADE'))

materialImportado = findEvenText(material_importado)
materialImportado.insert(0, ('MATERIAL IMPORTADO', 'QUANTIDADE'))

valorTotal = findEvenText(valor)
valorTotal.insert(0, ('TODOS VALORES', 'QUANTIDADE'))

for i in valor:
    if (i.endswith('EM GARANTIA')) is True:
        valor_garantia.append(i)
valorGarantia = findEvenText(valor_garantia)
valorGarantia.insert(0, ('VALOR EM GARANTIA', 'QUANTIDADE'))

for i in valor:
    if (i.endswith('SUCATADO')) is True:
        valor_sucatado.append(i)
valorSucatado = findEvenText(valor_sucatado)
valorSucatado.insert(0, ('VALOR EM SUCATA', 'QUANTIDADE'))

for i in valor:
    if (i.endswith('DE CONSERTO')) is True:
        valor_conserto.append(i)
valorConserto = findEvenText(valor_conserto)
valorConserto.insert(0, ('VALOR EM CONSERTO', 'QUANTIDADE'))

for i in df1['Nota']:
    qm.append(i)
qmQuant = findEvenText(qm)
qmQuant.insert(0, ('QM', 'QUANTIDADE'))

writeExcel(materiaisEnviadosConserto)
writeExcel(fornecedorOrdenado)
writeExcel(NFs)
writeExcel(materiaisSemDefeito)
writeExcel(materialAprovado)
writeExcel(materialAvariado)
writeExcel(materialDefeito)
writeExcel(materialLacreRompido)
writeExcel(materialImportado)
writeExcel(valorTotal)
writeExcel(valorGarantia)
writeExcel(valorSucatado)
writeExcel(valorConserto)
writeExcel(qmQuant)
