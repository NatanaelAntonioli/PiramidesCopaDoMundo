# --------------- DEPENDÊNCIAS ---------------

from xlutils.copy import copy
from xlrd import open_workbook
from random import random
import pandas as pd

# --------------- FUNÇÕES ---------------

# Mapeia letra para número

def letra(letra):
    return (ord(letra)) - 97

# Retorna o valor de uma célula

def celula(coluna, linha):
    return format(sh.cell_value(rowx=linha-1, colx=letra(coluna)))

# Escreve numa célula

def escreve(coluna, linha, texto):
    escrever_planilha.write(linha-1,letra(coluna), texto)

# --------------- CONFIGURAÇÕES ---------------

total_testes = 100

# --------------- ABERTURA DA PLANILHA ---------------

ler = open_workbook('aleatorio.xls',formatting_info=True)

ler_planilha = ler.sheet_by_index(0)

escrever = copy(ler)

escrever_planilha = escrever.get_sheet(0)


for j in range (0, total_testes):

    for i in range (2, 25):

        escreve ( 'a', i, random())

    nome = 'iteracao_' + str(j) + '.xls'

    escrever.save(nome)

    lido_pandas = pd.read_excel(nome)

    lido_pandas.sort_values(by=['Code'], inplace=True)

    lido_pandas.to_excel('iteracao_ordenada' + str(j)+ '.xls')

    







