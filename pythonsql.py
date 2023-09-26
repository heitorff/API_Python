import pymssql as sql
import pandas as pd
import time
#import warnings

dados = pd.read_excel(r"C:\Users\heito\OneDrive\Área de Trabalho\Code7\teste.xlsx")

conexao = sql.connect('172.30.0.22\sql2008', 'aytybd', 'bdaytyfull', 'AYTY_COREDEVN3')

cursor = conexao.cursor()

inicio = time.time()
for index,row in dados.iterrows():
    sql = "INSERT INTO TESTE (DESCRICAO, VALOR) VALUES (%s, %s)"
    val = (row['DESCRICAO'],row['VALOR'])
    cursor.execute(sql, val)
    conexao.commit()

final = time.time()

conexao.close()

print("Dados inseridos com sucesso no SQL")
print('Tempo de Processamento:', int(final - inicio), 'segundos')




#mailing = pd.read_sql_query('select nm_mailing from mailing where id_mailing = 1', conexao)

#print(mailing)