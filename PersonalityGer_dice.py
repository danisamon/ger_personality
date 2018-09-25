
# encoding: utf-8

import xlrd
import xlwt
from sklearn import tree
from sklearn.model_selection import cross_val_score
from imblearn.over_sampling import SMOTE
import numpy as np
from sklearn.model_selection import cross_val_predict

classificadores = {}
atributos = [18, 19, 20, 22, 23, 30, 32, 33, 36, 40]


classificadores[18]= None
classificadores[19]= None
classificadores[20]= None
classificadores[22]= None
classificadores[23]= None
classificadores[30]= None
classificadores[32]= None
classificadores[33]= None
classificadores[36]= None
classificadores[40]= None



caracteristicas= [
[1, 2, 0, 5, 5, 5, 0, 1, 3, 0, 1, 124, 63, 105, 17, 8, 54, 4, 5, 2, 7, 33, 18, 24, 4, 22, 49, 1, 0, 0, 0, 2, 17, 2, 1, 5, 0, 3],
[2, 5, 0, 2, 2, 5, 0, 4, 5, 5, 4, 128, 40, 6, 13, 11, 42, 3, 0, 2, 0, 5, 6, 11, 9, 37, 12, 2, 0, 33, 26, 2, 0, 18, 3, 12, 0, 2],
[3, 4, 0, 4, 4, 4, 0, 2, 0, 5, 2, 116, 44, 81, 9, 51, 20, 5, 0, 1, 2, 14, 6, 3, 7, 51, 46, 11, 0, 3, 0, 5, 11, 12, 0, 1, 9, 4],
[4, 2, 0, 2, 2, 2, 0, 2, 2, 0, 2, 126, 71, 26, 9, 38, 84, 6, 0, 8, 0, 2, 24, 14, 24, 70, 38, 11, 27, 0, 11, 14, 0, 15, 24, 3, 9, 2],
[5, 2, 0, 5, 5, 5, 0, 5, 1, 0, 5, 119, 67, 30, 11, 12, 81, 10, 1, 0, 34, 3, 4, 16, 6, 80, 9, 11, 8, 0, 0, 3, 0, 26, 3, 5, 0, 0],
[6, 5, 0, 2, 2, 5, 0, 1, 4, 0, 2, 123, 54, 70, 3, 16, 75, 5, 0, 2, 2, 2, 5, 6, 3, 18, 29, 0, 0, 0, 0, 0, 9, 0, 2, 13, 0, 1],
[7, 4, 0, 4, 4, 4, 0, 4, 5, 5, 4, 126, 48, 42, 11, 11, 85, 5, 0, 12, 9, 2, 6, 6, 12, 42, 4, 5, 0, 38, 0, 4, 0, 30, 2, 6, 0, 2],
[8, 2, 0, 2, 2, 2, 0, 2, 2, 1, 2, 127, 53, 102, 6, 13, 112, 8, 0, 9, 2, 0, 8, 9, 8, 45, 37, 7, 0, 3, 2, 3, 10, 3, 1, 3, 0, 2],
[9, 2, 0, 5, 5, 3, 0, 3, 3, 4, 0, 126, 39, 106, 20, 13, 29, 1, 0, 1, 0, 11, 5, 18, 7, 13, 27, 4, 0, 19, 0, 2, 0, 10, 10, 0, 0, 10],
[10, 5, 0, 2, 2, 3, 0, 4, 5, 1, 3, 125, 48, 13, 4, 10, 22, 1, 53, 0, 26, 0, 2, 8, 3, 38, 4, 53, 4, 0, 0, 6, 0, 33, 1, 2, 0, 1],
[11, 4, 0, 4, 4, 2, 0, 4, 2, 4, 4, 127, 48, 101, 16, 8, 36, 0, 40, 1, 1, 6, 6, 11, 2, 13, 47, 9, 0, 56, 0, 7, 0, 6, 0, 3, 0, 5],
[12, 2, 0, 2, 2, 1, 0, 2, 1, 0, 2, 125, 39, 13, 7, 41, 39, 2, 0, 7, 0, 8, 11, 16, 15, 125, 82, 4, 0, 0, 5, 8, 0, 15, 0, 7, 20, 2]
]


#lista geral de preferência de uso de cada atributo
#P = [18,20,19,23,33,32,22]

workbook = xlrd.open_workbook('b5-ref-dividido.xlsx')
worksheet = workbook.sheet_by_index(0)
worksheettest = workbook.sheet_by_index(1)
rowtitulo = worksheettest.row_values(0)


#percorrer por atributo menos sexo
for j in range (1,10):
	X = []
	Y = []
	
	
	for row_num in range(1, worksheet.nrows):
		pers = []
		aux = []

		row = worksheet.row_values(row_num)
		pers.append(float(row[4]))
		pers.append(float(row[5]))
		pers.append(float(row[6]))
		pers.append(float(row[7]))
		pers.append(float(row[8]))
		
		aux = caracteristicas[int(row[14])-1] + pers
		

		#Pega o vetor de características de acordo com o valor da imagem

		X.append(aux)

		if row_num == 1:
			print(X)

		if row[atributos[j]]:
			Y.append(1)
		else:
			Y.append(0)

	#balancear os dados
	sm = SMOTE(random_state=42)
	X_res, Y_res = sm.fit_sample(X, Y)			
	#print('Atributo %d' % (atributos[j]))
		
	classificadores[atributos[j]]= tree.DecisionTreeClassifier()
	
	classificadores[atributos[j]].fit(X_res,Y_res)
	
	
#TESTE
workbookw = xlwt.Workbook()
worksheetw = workbookw.add_sheet(u'DICE')


worksheetw.write(0,4,u'Dice')
worksheetw.write(0,0,u'Imagem')
worksheetw.write(0,5,u'Id Participante')
worksheetw.write(0,1,u'Participante')
worksheetw.write(0,2,u'Sistema')
worksheetw.write(0,3,u'Intersecção')

for row_num in range(1, worksheettest.nrows):
	id = int(row[14]) - 1
	idpart = int(row[0]) 

	row = worksheettest.row_values(row_num)
	XT=[]
	L = []
	Ld= []
	
	pers = []
	aux = []

	
	pers.append(float(row[4]))
	pers.append(float(row[5]))
	pers.append(float(row[6]))
	pers.append(float(row[7]))
	pers.append(float(row[8]))
		
	aux = caracteristicas[int(row[14])-1] + pers
		

		#Pega o vetor de características de acordo com o valor da imagem 
	XT.append(aux)

	print(XT)
	#adiciono sexo automaticamente
	L.append(rowtitulo[atributos[0]])

	for j in range (1,10):
		
		if classificadores[atributos[j]].predict(XT) == 1:
			#L[rowtitulo[atributos[j]]] = row[atributos[j]]
			L.append(rowtitulo[atributos[j]])

	for i in range(18,45):
		if row[i] != '' :
			#Ld[rowtitulo[i]] = row[i]
			Ld.append(rowtitulo[i])
	print(Ld)
	#print(L)
	intersection = set(Ld).intersection(L)
    #print(intersection)
	dice = (2 * len(intersection)) / (len(Ld) + len(L))
    #print(dice)

	worksheetw.write(row_num, 4, dice)
	worksheetw.write(row_num, 0,   (id + 1))
	worksheetw.write(row_num, 5,   idpart)
	worksheetw.write(row_num, 1, str(Ld).strip("{}"))
	worksheetw.write(row_num, 2, str(L).strip('{}'))
	worksheetw.write(row_num, 3, ', '.join(intersection))
    

workbookw.save('dice_PersonalityGer.xls')
#TESTE


