import pandas as pd
import pandasql as psql
import sqlite3
import pyexcel as p
import xlrd


from datetime import datetime


#tentei encontrar bibliotecas para altertar o arq   uivo xls para xlxs, pyexcel pyexcel-cli e subbibliotecas nao funcionaram
#xls2xlsx tambem nao funcionaram
#processo de automatizacao para estas bibliotecas nao funcionaram, o idel neste caso e abrir o documento e salvar em xls (nao reniomear pois nao funciona)


#Script deve ser rodado no mesmo arquivo em que esta a planilha excel
create_at = datetime.now()

oleodf=  pd.read_excel("./vendas-combustiveis-m3.xlsx",sheet_name="DPCache_m3")
oleodf2=  pd.read_excel("./vendas-combustiveis-m3.xlsx",sheet_name="DPCache_m3_2")
oleodf3= pd.concat([oleodf,oleodf2])

#limpeza e normalizaçao dos dados
#extracao de combuntiveis nao sao a base de petroleo
oleodf4=oleodf3.query('COMBUSTÍVEL!="GLP (m3)"')
oleodf4=oleodf4.query('COMBUSTÍVEL!="ETANOL HIDRATADO (m3)"')

#criaão do campo create at
oleodf4['CREATE_AT']=create_at

#preenchimento de dados para campo unit baseado no campo combustivel e extraçao da unidade do campo combustivel
unit = []
combustivel = []

for i in oleodf4['COMBUSTÍVEL'].index:
    unit.append(oleodf4['COMBUSTÍVEL'][i][oleodf4['COMBUSTÍVEL'][i].find('(')+1:oleodf4['COMBUSTÍVEL'][i].find(')')])
    combustivel.append(oleodf4['COMBUSTÍVEL'][i][0:oleodf4['COMBUSTÍVEL'][i].find('(')])

se=pd.Series(unit)
oleodf4['UNIDADE']=se.values
se=pd.Series(combustivel)
oleodf4['COMBUSTÍVEL']=se.values


#renomeando nomes de colunas
oleodf4.rename(columns = {'COMBUSTÍVEL':'PRODUTO'}, inplace=True)
oleodf4.rename(columns = {'ESTADO': 'UF'}, inplace=True)
oleodf4.rename(columns = {'Jan': 'Janeiro', 'Fev':'Fevereiro', 'Mar':'Marco', 'Abr':'Abril'}, inplace=True)
oleodf4.rename(columns = {'Mai':'Maio', 'Jun':'Junho', 'Jul':'Julho'}, inplace=True)
oleodf4.rename(columns = {'Ago':'Agosto', 'Set':'Setembro', 'Out':'Outubro', 'Nov':'Novembro', 'Dez':'Dezembro'}, inplace=True)

#Criaçao da coluna volume acumulados no mes
oleodf4['VOLUME']=oleodf4['Janeiro']+oleodf4['Fevereiro']+oleodf4['Marco']+oleodf4['Abril']+oleodf4['Maio']+oleodf4['Junho']+oleodf4['Julho']+oleodf4['Agosto']+oleodf4['Setembro']+oleodf4['Outubro']+oleodf4['Novembro']+oleodf4['Dezembro']

#separação dos combustiveis fosseis em diesel(todos e derivados(demais
oleodieseldf=psql.sqldf("select * from oleodf4 where COMBUSTÍVEL LIKE '%DIESEL%'")
derivadodf=psql.sqldf("select * from oleodf4 where COMBUSTÍVEL NOTLIKE '%DIESEL%'")


oleodieseldf=psql.sqldf("Select ANO as year_month, UF, PRODUTO, UNIDADE,VOLUME, CREATE_AT  from oleodf4 where COMBUSTÍVEL LIKE '%DIESEL%'")
derivadodf=psql.sqldf("Select ANO as year_month, UF, PRODUTO, UNIDADE,VOLUME, CREATE_AT  from oleodf4 where COMBUSTÍVEL NOTLIKE '%DIESEL%'")

#armazenar em base de dados e excel(copia de seguranca

#CRIAR CONEXAO
con = sqlite3.connect("./combustivel.sqlite")

oleodieseldf.to_sql("T_Oleo_Diesel", con, if_exists="replace")
derivadodf.to_sql("T_Derivados", con, if_exists="replace")
con.close()

#criando copia em arquivos excel
oleodieseldf.to_excel("Oleo_Diesel")
derivadodf.to_excel("derivados")

#finalizaçao
print ("Processo finalizado ")


