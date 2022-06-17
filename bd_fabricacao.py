import openpyxl
import os
import pandas as pd

pasta_arquivos_originais = r'C:\Users\PedroJesusPalmont\Palmont\Controle de Obras - General\Arquivos Base\Bancos de dados\Histogramas\ORCADO'
pasta_arquivos_copiados = r'C:\Users\PedroJesusPalmont\Palmont\Controle de Obras - General\Arquivos Base\Bancos de dados\Histogramas\Copia de orcado'


def adicionar_fabricacao():
    for arquivo in os.listdir(pasta_arquivos_originais):
        caminho_arquivo = fr'{pasta_arquivos_originais}\{arquivo}'

        wb = openpyxl.load_workbook(caminho_arquivo)
        planilha = wb['Banco de dados']
        adicionar_valores(planilha=planilha)
        planilha_copiada = pasta_arquivos_copiados + '\\' + arquivo
        wb.save(planilha_copiada)


def adicionar_valores(planilha):
    primeira_linha = 4562

    # valor i utilizado nas condições de letras
    i = 0
    for linha in range(primeira_linha, primeira_linha+399, 21):
        for c in range(0, 21):
            linha_atual = linha + c

            # celula A
            A = r"='Histograma MOI Geral'!$U$2"
            planilha[f'A{linha_atual}'] = A

            # celula B
            lista = [7, 8, 9, 11, 12, 13, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29]
            B = f"='Histograma Fabricação'!A{lista[c]}"
            planilha[f'B{linha_atual}'] = B

            # celula C
            if c <= 2:
                C = "='Histograma Fabricação'!A6"
            elif c <= 5:
                C = "='Histograma Fabricação'!A10"
            else:
                C = "='Histograma Fabricação'!A14"
            planilha[f'C{linha_atual}'] = C

            # celula D
            lista_letras = ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T']
            numero = lista_letras[i]
            D = f"='Histograma Fabricação'!{numero}5"
            planilha[f'D{linha_atual}'] = D

            # celula E
            numero_e = lista[c]
            valor_e = int((linha-primeira_linha)/21)
            letra_e = lista_letras[valor_e]

            E = f"='Histograma Fabricação'!{letra_e}{numero_e}"
            planilha[f'E{linha_atual}'] = E

            # celula F
            planilha[f'F{linha_atual}'] = 'FABRICACAO'
        i += 1


if __name__ == '__main__':
    adicionar_fabricacao()
