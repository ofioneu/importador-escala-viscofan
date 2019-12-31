import pandas as pd
import xlrd
from pandas import read_excel


a = pd.read_excel('TESTE VISCOFAN 31122019.xlsx') # abre o arquivo xlsx

frame = pd.DataFrame(a) # cria um dataframe geral do arquivo



t1e = frame.loc[0, 'TURNO 1 E']
t1s = frame.loc[0, 'TURNO 1 S']
t2e = frame.loc[0, 'TURNO 2 E']
t2s = frame.loc[0, 'TURNO 2 S']
t3e = frame.loc[0, 'TURNO 3 E']
t3s = frame.loc[0, 'TURNO 3 S']

ida = frame.loc[0, 'ID A'].astype(int)
idb = frame.loc[0, 'ID B'].astype(int)
idc = frame.loc[0, 'ID C'].astype(int)
idd = frame.loc[0, 'ID D'].astype(int)
ide = frame.loc[0, 'ID E'].astype(int)

ano = frame.loc[0, 'ANO'].astype(int)
mes = frame.loc[0, 'MÊS'].astype(int)

dia = 0
for i in frame['ESCALA A'].dropna():
    dia=dia+1
    if (i == 1):
        l = open ('escala A mes-'+str(mes)+'.txt', 'a')
        l.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"'"",'Exceção criada em 24/01/2019',"+str(t1e)+",'Exceção',"+str(t1s)+","+str(ida)+");"+'\n')
        l.close ()
    
    if (i == 2):
        l = open ('escala A mes-'+str(mes)+'.txt', 'a')
        l.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"'"",'Exceção criada em 24/01/2019',"+str(t2e)+",'Exceção',"+str(t2s)+","+str(ida)+");"+'\n')
        l.close ()
    
    if (i == 3):
        l = open ('escala A mes-'+str(mes)+'.txt', 'a')
        l.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"'"",'Exceção criada em 24/01/2019',"+str(t3e)+",'Exceção',"+str(t3s)+","+str(ida)+");"+'\n')
        l.close ()
print('Arquivo criado com sucesso!!')

dia=0
for i in frame['ESCALA B'].dropna():    
    dia=dia+1
    if (i == 1):
        l = open ('escala B mes-'+str(mes)+'.txt', 'a')
        l.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"'"",'Exceção criada em 31/12/2019',"+str(t1e)+",'Exceção',"+str(t1s)+","+str(idb)+");"+'\n')
        l.close ()
    
    if (i == 2):
        l = open ('escala B mes-'+str(mes)+'.txt', 'a')
        l.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"'"",'Exceção criada em 31/12/2019',"+str(t2e)+",'Exceção',"+str(t2s)+","+str(idb)+");"+'\n')
        l.close ()
    
    if (i == 3):
        l = open ('escala B mes-'+str(mes)+'.txt', 'a')
        l.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"'"",'Exceção criada em 31/12/2019',"+str(t3e)+",'Exceção',"+str(t3s)+","+str(idb)+");"+'\n')
        l.close ()
print('Arquivo criado com sucesso!!')

dia=0
for i in frame['ESCALA C'].dropna():    
    dia=dia+1
    if (i == 1):
        l = open ('escala C mes-'+str(mes)+'.txt', 'a')
        l.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"'"",'Exceção criada em 24/01/2019',"+str(t1e)+",'Exceção',"+str(t1s)+","+str(idc)+");"+'\n')
        l.close ()
    
    if (i == 2):
        l = open ('escala C mes-'+str(mes)+'.txt', 'a')
        l.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"'"",'Exceção criada em 24/01/2019',"+str(t2e)+",'Exceção',"+str(t2s)+","+str(idc)+");"+'\n')
        l.close ()
    
    if (i == 3):
        l = open ('escala C mes-'+str(mes)+'.txt', 'a')
        l.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"'"",'Exceção criada em 24/01/2019',"+str(t3e)+",'Exceção',"+str(t3s)+","+str(idc)+");"+'\n')
        l.close ()
print('Arquivo criado com sucesso!!')

dia=0
for i in frame['ESCALA D'].dropna():    
    dia=dia+1
    if (i == 1):
        l = open ('escala D mes-'+str(mes)+'.txt', 'a')
        l.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"'"",'Exceção criada em 24/01/2019',"+str(t1e)+",'Exceção',"+str(t1s)+","+str(idd)+");"+'\n')
        l.close ()
    
    if (i == 2):
        l = open ('escala D mes-'+str(mes)+'.txt', 'a')
        l.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"'"",'Exceção criada em 24/01/2019',"+str(t2e)+",'Exceção',"+str(t2s)+","+str(idd)+");"+'\n')
        l.close ()
    
    if (i == 3):
        l = open ('escala D mes-'+str(mes)+'.txt', 'a')
        l.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"'"",'Exceção criada em 24/01/2019',"+str(t3e)+",'Exceção',"+str(t3s)+","+str(idd)+");"+'\n')
        l.close ()
print('Arquivo criado com sucesso!!')

dia=0
for i in frame['ESCALA E'].dropna():
    dia=dia+1
    if (i == 1):
        l = open ('escala E mes-'+str(mes)+'.txt', 'a')
        l.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"'"",'Exceção criada em 24/01/2019',"+str(t1e)+",'Exceção',"+str(t1s)+","+str(ide)+");"+'\n')
        l.close ()
    
    if (i == 2):
        l = open ('escala E mes-'+str(mes)+'.txt', 'a')
        l.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"'"",'Exceção criada em 24/01/2019',"+str(t2e)+",'Exceção',"+str(t2s)+","+str(ide)+");"+'\n')
        l.close ()
    
    if (i == 3):
        l = open ('escala E mes-'+str(mes)+'.txt', 'a')
        l.write ("INSERT INTO dia (id,data,descricao,entrada,nome,saida,escala) VALUES (nextval('dia_seq'),"+"'"+str(ano)+"-"+str(mes)+"-"+str(dia)+"'"",'Exceção criada em 24/01/2019',"+str(t3e)+",'Exceção',"+str(t3s)+","+str(ide)+");"+'\n')
        l.close ()
print('Arquivo criado com sucesso!!')



