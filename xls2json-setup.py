# -*- coding: utf-8 -*-
import openpyxl
import io
import json
import sys


class Setup:

 	def __init__(self,ID,versao, data, produtividadePF, servicos):
	 	        
	        self.ID=ID
	        self.versao=versao
	        self.data=data
	        self.produtividadePF=produtividadePF
	        self.servicos=servicos


def newItem(data, sigla):
    if sigla in data:
		return False
    else:
		return True;
def setPath():
    currentPath = sys.path[0]
    if currentPath.find('/')!=-1: # found
        currentPath = currentPath + "/"
    else:
        currentPath = currentPath + "\\"

    return currentPath

myPath =''
myPath = setPath();

wb = openpyxl.load_workbook('Set-up.xlsx')
mysheet = ""
sheet = wb.get_sheet_by_name('Planos')

setupID = sheet['B1'].value.strip()
versao = sheet['D1'].value.strip()
dataVersao = sheet['F1'].value
produtividadePF = sheet['H1'].value


line = 3
mycell = 'A' + str(line)
servicos = []
itemservico =[]

descricaoservico =''

line = 3
mycell = 'B' + str(line)
descricaoitemservico = sheet['B' + str(line)].value.strip()
#siglaproduto = sheet['E' + str(line)].value.strip()
descricaoservico = sheet['A' + str(line)].value.strip()
produtos=[]

while sheet[mycell].value:
	
	#if (siglaproduto != sheet['E' + str(line)].value.strip()) and (descricaoitemservico == sheet['B' + str(line)].value.strip()) and (descricaoservico == sheet['A' + str(line)].value.strip()):	

	#if newItem(produtos, sheet['E' + str(line)].value.strip()):
	if (descricaoitemservico == sheet['B' + str(line)].value.strip()) and (descricaoservico == sheet['A' + str(line)].value.strip()):
		p = {"sigla":sheet['E' + str(line)].value.strip(), "descricao":sheet['F' + str(line)].value.strip(), "tipo":sheet['G' + str(line)].value.strip(), "homologador":sheet['H' + str(line)].value.strip()}
		produtos.append(p)

	if (descricaoitemservico != sheet['B' + str(line)].value.strip()): #and (descricaoservico == sheet['A' + str(line)].value.strip()):
		descricaoitemservico = sheet['B' + str(line - 1)].value.strip()
		localitemservico = {"descricao":descricaoitemservico, "produtos":produtos}
		itemservico.append(localitemservico)
		produtos=[]
		descricaoitemservico = sheet['B' + str(line)].value.strip()
		p = {"sigla":sheet['E' + str(line)].value.strip(), "descricao":sheet['F' + str(line)].value.strip(), "tipo":sheet['G' + str(line)].value.strip(), "homologador":sheet['H' + str(line)].value.strip()}
		produtos.append(p)

	if descricaoservico != sheet['A' + str(line)].value.strip():
		s = {"descricao":sheet['A' + str(line -1)].value.strip(), "itemDeServico":itemservico}
		servicos.append(s)
		itemservico =[]
		produtos=[]
		p = {"sigla":sheet['E' + str(line)].value.strip(), "descricao":sheet['F' + str(line)].value.strip(), "tipo":sheet['G' + str(line)].value.strip(), "homologador":sheet['H' + str(line)].value.strip()}
		produtos.append(p)
		descricaoservico = sheet['A' + str(line)].value.strip()

	line += 1
	mycell = 'B' + str(line)

#print itemservico
setup=Setup(setupID,versao, dataVersao, produtividadePF, servicos)
#print servico
#print json.dumps(vars(setup),sort_keys=True, indent=4)  

with io.open('setup.json', 'w', encoding='utf8') as json_file:
	data = json.dumps(vars(setup),sort_keys=True, indent=4, ensure_ascii=False)
	json_file.write(unicode(data))

#fileContenent =  json.dumps(vars(setup),sort_keys=True, indent=4, ensure_ascii=False).encode('utf8')

#with io.open(myPath + 'setup.json','w',encoding='utf8') as f:
#	f.write(fileContenent)

#with io.open('setup.json', 'w', encoding='utf8') as json_file:
#	json.dumps(vars(setup),sort_keys=True, indent=4)