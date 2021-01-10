## Esse é um código usado no vídeo Desmistificando: Pirâmide da Copa do Mundo,
## que provavelmente sairá em janeiro de 2021. Se você chegou nesse diretório 
## depois dessa data, procure o vídeo no canal Fábrica de Noobs para encontrar documentação.

# --------------- PRELIMINARES ------------------------

#Dependências

import xlrd #Precisa de pip
import math

# Lê a planilha

plan = xlrd.open_workbook("copas.xls")
sh = plan.sheet_by_index(0)

# Mapeia letra para número

def letra(letra):
    return (ord(letra)) - 97

# Retorna o valor de uma célula

def celula(coluna, linha):
    return format(sh.cell_value(rowx=linha-1, colx=letra(coluna)))

# --------------- RESTRIÇÕES     ------------------------

usarrestricoes = 0 # Ativa as restrições
restricao_raio = 9 # Define o raio de busca, ou seja, quantos "degraus" a pirâmide terá. Coloque 9 para uma comparação justa com a Copa do Mundo.

valor_limiar = 3 # Notifica ao final caso o total de acertos seja maior ou igual

celula_times = 'b'
celula_anos = 'c'

# -------------- LISTAS PARA ESTATÍSTICA FINAL --------------

anos = []
acertos = []
limiar = []
totais = [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]

# --------------- CÓDIGO PRINCIPAL ---------------------

for j in range (0,131):

    match = 0
    meio = j

    if (usarrestricoes == 1): 
        raio = restricao_raio + 1
    else:
        raio = 200
    
    if (celula ('b', meio) != '' and celula ('b', meio) != '.'):

        print ("")
        print ("------------- TESTANDO PARA CENTRO EM " + celula('c', meio) + ' ( ' + celula ('b', meio) + ' ) ' + "---------- SOMA: " + str(2 * float(celula ('b', meio))))
        print ("")
        for i in range (1, raio):

            if (celula('c', meio) == ''):
                break

            if ((celula('c', meio + i)) != '') and ((celula('c', meio - i)) != ''):
            
                if ((celula('c', meio + i)) == (celula('c', meio - i))):

                    if ((celula('c', meio + i) != '')):

                        match = match + 1
                        print (celula('c', meio + i) + '(' + celula ('b', meio +i) + ')' + 'vs ' + celula('c', meio - i) + '(' + celula ('b', meio - i) + ')' + " --> ACERTO!")
                else:
                        print (celula('c', meio + i) + '(' + celula ('b', meio +i) + ')' + 'vs ' + celula('c', meio - i) + '(' + celula ('b', meio - i) + ')')
                        
        print ("Total de acertos:", match)

        totais[match]= totais[match]+1
            

        anos.append(celula('c', meio) + '(' + celula ('b', meio) + ')')
        acertos.append(match)

        if (match >= valor_limiar):
            limiar.append(' --> LIMIAR!')
        else:
            limiar.append('')
        
    
print ("")
print ("-------------- TESTES DE CADA PIRÂMIDE ----------------- ")
print ("")

for x in range(len(anos)):
    print (anos[x] + "com " + str(acertos[x]) + " acertos" + limiar[x])

print ("")
print ("-------------- TOTAL DE PIRÂMIDES ----------------- ")


print ("")

if (usarrestricoes == 1):
    for x in range(0, raio):

        formatado = "{:.0f}".format(x/(restricao_raio)*100)

        print (str(totais[x]) + " pirâmides com " + str(x) + " acertos. (" + str(formatado) + "%)")
