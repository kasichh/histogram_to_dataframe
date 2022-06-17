import openpyxl


def apply_formulas():
    """
    Substituir coluna de MOI por MOD
    :return: Null
    """
    wb = openpyxl.load_workbook('histogramapadrao.xlsx')
    planilha = wb['BD (2)']
    for c in range(2, planilha.max_row):
        for d in ['A', 'B', 'C', 'D']:
            try:
                numero = c + 1900
                formula = planilha[f'{d}{c}'].value
                planilha[f'{d}{numero}'] = formula.replace('Histograma MOI Geral', 'Histograma MOD Geral')
            except:
                pass

    wb.save('saida.xlsx')


if __name__ == '__main__':
    apply_formulas()
