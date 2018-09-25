# encoding: utf-8

import xlrd
import xlwt


def eliminaDistraidor(C, valor, p):
    tam = len(C)
    Caux = []
    for i in range(0, tam):
        if C[i][p] == valor:
            Caux.append(C[i])
    return Caux

def verificaDistraidores(C, valor, p):
    tam = len(C)
    Caux = []
    for i in range(0, tam):
        if C[i][p] != valor:
            Caux.append(C[i])
    return Caux

#workbook = xlrd.open_workbook('b5-clean-images.xlsx')
workbook = xlrd.open_workbook('b5-clean-images.xlsx')
worksheet = workbook.sheet_by_index(0)
L = [[], [], [], [], [],[], [], [],[], [],[], []]
C = [[], [], [], [], []]
t = []
#Preferencia Geral
P = [5, 7, 10, 6, 19, 20, 9, 27, 23, 17, 8, 21, 16, 12, 18, 15, 14, 29, 25, 11, 26, 13, 28, 30, 24, 22, 31]

#Preferencia de isExt
#P = [5, 7, 19, 10, 6, 20, 27, 23, 17, 9, 18, 8, 21, 16, 15, 12, 14, 13, 29, 24, 26, 25, 31, 22, 28, 30, 11]
#Preferencia de isNotExt
#P = [5, 6, 10, 7, 19, 20, 9, 23, 27, 8, 17, 11, 12, 21, 14, 15, 18, 16, 26, 30, 13, 28, 29, 25, 22, 24, 31]

#Preferencia de isAgr
#P = [5, 6, 7, 10, 19, 20, 9, 8, 27, 17, 21, 16, 12, 15, 18, 23, 14, 31, 29, 11, 13, 25, 30, 26, 22, 24, 28]
#Preferencia de isNotAgr
#P = [5, 10, 6, 7, 19, 20, 9, 23, 27, 8, 14, 17, 18, 21, 12, 16, 29, 11, 13, 15, 22, 25, 30, 28, 26, 24, 31]

#Preferencia de isCon
#P = [5, 7, 6, 10, 19, 20, 9, 27, 8, 23, 21, 16, 12, 14, 15, 17, 18, 25, 29, 22, 11, 31, 30, 13, 26, 24, 28]
#Preferencia de isNotCon
#P = [10, 7, 6, 19, 20, 9, 23, 8, 21, 12, 17, 27, 28, 14, 15, 16, 18, 13, 29, 30, 11, 24, 26, 25, 22, 31]

#Preferencia de isNeu
#P = [5, 7, 10, 6, 19, 20, 9, 17, 16, 27, 18, 23, 14, 21, 8, 15, 12, 29, 11, 13, 25, 24, 26, 31, 22, 30, 28]
#Preferencia de isNotNeu
#P = [5, 10, 7, 19, 6, 20, 9, 17, 8, 27, 12, 21, 23, 14, 16, 18, 15, 11, 29, 30, 13, 25, 26, 24, 31, 28, 22]

#Preferencia de isOpe
#P = [5, 7, 10, 6, 19, 20, 21, 23, 27, 9, 17, 12, 15, 18, 25, 8, 14, 28, 16, 22, 30, 24, 29, 31, 11, 13, 26]
#Preferencia de isNotOpe
#P = [5, 10, 19, 7, 6, 20, 9, 27, 17, 8, 23, 16, 18, 21, 12, 14, 15, 29, 26, 11, 13, 25, 31, 28, 30, 24, 22]


aux = 0
auxi = 1
auxf = 7
rowtitulo = worksheet.row_values(0)

for x in range (0,12):

    aux = 0
    C = [[], [], [], [], []]

    for row_num in range(auxi,auxf):

        row = worksheet.row_values(row_num)

        if row[3] == 'target':
            t = row
        else:
            C[aux] = row
            aux += 1

    for i in range(0, 27):
        if verificaDistraidores(C, t[P[i]], P[i]):
            #L[x][rowtitulo[P[i]]] = t[P[i]]
            L[x].append(rowtitulo[P[i]])
            C = eliminaDistraidor(C, t[P[i]], P[i])
        if not C:
            break

    auxi = auxf
    auxf += 6

#workbookd = xlrd.open_workbook('b5-clean-descriptions.xlsx')
workbookd = xlrd.open_workbook('b5-ref-dividido.xlsx')
worksheetd = workbookd.sheet_by_index(1)
rowtitulo = worksheetd.row_values(0)
workbookw = xlwt.Workbook()
worksheetw = workbookw.add_sheet(u'DICE')


worksheetw.write(0,4,u'Dice')
worksheetw.write(0,0,u'Imagem')
worksheetw.write(0,1,u'Participante')
worksheetw.write(0,2,u'Sistema')
worksheetw.write(0,3,u'Intersecção')
worksheetw.write(0,5,u'Id Participante')


for row_numd in range(1,worksheetd.nrows):
    Ld = []
    row = worksheetd.row_values(row_numd)
    id = int(row[14]) - 1
    idpart = int(row[0]) 

    #for i in range(17,44):
     #   if row[i] != '' :
      #      Ld[rowtitulo[i-12]] = row[i]

    for i in range(18,45):
        if row[i] != '' :

            Ld.append(rowtitulo[i])

    print(Ld)
    print(L[id])
    intersection = set(Ld).intersection(L[id])
    #print(intersection)
    dice = (2 * len(intersection)) / (len(Ld) + len(L[id]))
    #print(dice)

    worksheetw.write(row_numd, 4, dice)
    worksheetw.write(row_numd, 0,   (id + 1))
    worksheetw.write(row_numd, 1, str(Ld).strip("{}"))
    worksheetw.write(row_numd, 2, str(L[id]).strip('{}'))
    worksheetw.write(row_numd, 3, ', '.join(intersection))
    worksheetw.write(row_numd, 5,   idpart)
    

workbookw.save('dice_AlgoritmoIncremental.xls')


