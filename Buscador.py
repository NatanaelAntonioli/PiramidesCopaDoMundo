# --------------- PRELIMINARES ------------------------

#Dependências

import xlrd #Precisa de pip
import xlwt #Precisa de pip
import math


# Mapeia letra para número

def letra(letra):
    return (ord(letra)) - 97

# Retorna o valor de uma célula

def celula(coluna, linha):
    return format(sh.cell_value(rowx=linha-1, colx=letra(coluna)))

# --------------- RESTRIÇÕES     ------------------------

usarrestricoes = 1 # Ativa as restrições
restricao_raio = 9 # Define o raio de busca, ou seja, quantos "degraus" a pirâmide terá. Coloque 9 para uma comparação justa com a Copa do Mundo.

valor_limiar = 3 # Notifica ao final caso o total de acertos seja maior ou igual

total_iteracoes = 10

# -------------- LISTAS PARA ESTATÍSTICA FINAL --------------

totais_finais = [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]

# --------------- CÓDIGO PRINCIPAL ---------------------


## Para cada conjunto

for contador_iteracoes in range (0, total_iteracoes):

    anos = []
    acertos = []
    limiar = []
    totais = [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]

    nome = 'iteracao_ordenada' + str(contador_iteracoes) + '.xls'
     
    plan = xlrd.open_workbook(nome)
    
    sh = plan.sheet_by_index(0)


    for j in range (2,24):

        match = 0
        meio = j

        if (usarrestricoes == 1): 
            raio = restricao_raio + 1
        else:
            raio = 200
        
        if (celula ('c', meio) != '' and celula ('c', meio) != '.'):

            #print ("")
            #print ("------------- TESTANDO PARA CENTRO EM " + celula('d', meio) + ' ( ' + celula ('c', meio) + ' ) ' + "---------- SOMA: " + str(2 * float(celula ('c', meio))))
            #print ("")
            for i in range (1, raio):

                if (celula('d', meio) == ''):
                    break

                if ((celula('d', meio + i)) != '') and ((celula('d', meio - i)) != ''):
                
                    if ((celula('d', meio + i)) == (celula('d', meio - i))):

                        if ((celula('d', meio + i) != '')):

                            match = match + 1
                            #print (celula('d', meio + i) + '(' + celula ('c', meio +i) + ')' + 'vs ' + celula('d', meio - i) + '(' + celula ('c', meio - i) + ')' + " --> ACERTO!")
                    else:
                            a = 1
                            #print (celula('d', meio + i) + '(' + celula ('c', meio +i) + ')' + 'vs ' + celula('d', meio - i) + '(' + celula ('c', meio - i) + ')')
                            
            #print ("Total de acertos:", match)

            totais[match]= totais[match]+1
                

            anos.append(celula('d', meio) + '(' + celula ('c', meio) + ')')
            acertos.append(match)

            if (match >= valor_limiar):
                limiar.append(' --> LIMIAR!')
            else:
                limiar.append('')
            
        
    print ("")
    print ("------------------ TESTANDO: " + nome + "---------------")
    print ("")

    for x in range(len(anos)):
        print (anos[x] + "com " + str(acertos[x]) + " acertos" + limiar[x])


    for x in range(0, 9):

        formatado = "{:.0f}".format(x/(restricao_raio)*100)

        if (totais[x] != 0):
            
            totais_finais[x] = totais_finais[x] + 1 


print ("")

print ("-------------- TOTAIS FINAIS DE PIRÂMIDES ----------------- ")


print ("")
    
for y in range (0,9):

    print (str(totais_finais[y]) + " de " + str(total_iteracoes) + " ordenações aleatórias têm ao menos uma pirâmide de " + str(y) + " acertos.")
