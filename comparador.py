from openpyxl import workbook, load_workbook;

mocca = load_workbook('indicadorProd.xlsx');

aba_ativa = mocca['Horimetro'];

for cell in aba_ativa['C']:
    if (cell.value != None):
        if(cell.value != "Data"):
            if(str(cell.value).replace(' 00:00:00', "") == '2023-03-07'):
                m1_H_inicial = aba_ativa[f'E{cell.row}'].value
                m1_H_final = aba_ativa[f'F{cell.row}'].value
                print(('{:.1f}'.format(m1_H_final-m1_H_inicial)))