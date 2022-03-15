import pyodbc
import pandas as pd
import openpyxl


servidor= input('Introduce el nombre del servidor: ')

server = servidor+'\SQLEXPRESS'
database = 'C:\MyBusinessDatabase\MyBusinessPOS2011.mdf' 
username = 'sa' 
password = '12345678'


try:
    cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
    cursor = cnxn.cursor()
    print('conexion exitosa')

except:
    print('no conecto')

try:
    
    dfup=pd.read_excel('ARTICULOS ROTACION 2.xlsx')
    print('lectura de excel Ok')
    
except Exception as FaltaXLSX:
    print(FaltaXLSX)

#-----------------------------se importa la columna Articulo de clavesadd en un df
query = "SELECT [Articulo] FROM clavesadd;"
dfArti = pd.read_sql(query, cnxn)
#-----------------------------------------comparar Articulo de clavesadd con viejo
for i in range(0,12):
    clave=str(dfArti.loc[i,'Articulo'])
    for j in range(0,1201):
        ArtiV=str(dfup.loc[j,'Viejo'])
        if clave==ArtiV:
            ArtiN=dfup.loc[j,'Nuevo']
            print(i,j,ArtiN,ArtiV) #---------------------------------actualizar prods
            cursor.execute("UPDATE Clavesadd SET Clave='"+ArtiN+"' WHERE ARTICULO='"+ArtiV+"'")


        
cnxn.commit()
cursor.close()
