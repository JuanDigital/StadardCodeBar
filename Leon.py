# -*- coding: utf-8 -*-
"""
Created on Wed Feb  2 14:59:37 2022

@author: Juan Carlos Rodríguez Pérez
"""

import pyodbc
import pandas as pd
i=0
j=0

#server = 'DESKTOP-92H6KQK\SQLEXPRESS'
#bd = 'C:\MyBusinessDatabase\MyBusinessPOS2011.mdf'
#usuario='sa'
#ºcontraseña='12345678'

#realizar conexion a BBDD

try:
    
    conexion=pyodbc.connect('DRIVER={SQL Server};SERVER=DESKTOP-92H6KQK\SQLEXPRESS;DATABASE=C:\MyBusinessDatabase\MyBusinessPOS2011.mdf;UID=sa;PWD=12345678')
    #print ('conexión exitosa!!!')
    cursor=conexion.cursor()
    #cursor.execute('SELECT *from prods;')
    #row=cursor.fetchone()
    #print(type(row))
except Exception as FueraDeLinea:
    print(FueraDeLinea)
    
 #lectura de archivo excel   
try:
    
    dfup=pd.read_excel('ARTICULOS ROTACION 2.xlsx')
    
except Exception as FaltaXLSX:
    print(FaltaXLSX)
#-----------------------------se importan la columna Articulo de prods en un df
#query = "SELECT [Articulo], [Descrip] FROM Prods;"
query = "SELECT [Articulo] FROM Prods;"
dfArti = pd.read_sql(query, conexion)
#print(dfArti.head)
'''
#------------------------------------------------------eliminar tabla clavesadd
cursor.execute("DELETE FROM Clavesadd") 

#---------------------------------------------Insert Dataframe into SQL Server:
for index, row in dfArti.iterrows():
     cursor.execute("INSERT INTO Clavesadd (Clave,Dato1,Usuario,usuFecha,usuHora,Dato2,Articulo,Unidad,Imagen) VALUES(?,0,'Sys','2022-02-07 00:00:00.000','13:00:00',0,0,0,0)", row.Articulo)
'''
#-----------------------------------------comparar Articulo de prods con viejo
for i in range(0,1377):
    clave=str(dfArti.loc[i,'Articulo'])
    for j in range(0,1201):
        ArtiV=str(dfup.loc[j,'Viejo'])
        if clave==ArtiV:
            ArtiN=dfup.loc[j,'Nuevo']
            print(i,j,ArtiN) #---------------------------------actualizar prods
            cursor.execute("UPDATE prods SET ARTICULO='"+ArtiN+"' WHERE ARTICULO='"+ArtiV+"'")
        #else:
         #   print()
'''    
clave=str(dfArti.loc[219,'Articulo'])            
ArtiV=str(dfup.loc[0,'Viejo'])
ArtiN=str(dfup.loc[0,'Nuevo'])


print(clave,ArtiV)
if clave<=ArtiV:
    print('ok')
else: print('nok')    
    #ArtiN=dfup.loc[i,'Nuevo']
    #print(ArtiV,ArtiN)

 

print("UPDATE prods SET ARTICULO='"+ArtiN+"' WHERE ARTICULO='"+ArtiV+"'")
cursor.execute("UPDATE prods SET ARTICULO='"+ArtiN+"' WHERE ARTICULO='"+ArtiV+"'")
'''

conexion.commit()
cursor.close()
