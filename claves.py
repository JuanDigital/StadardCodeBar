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
'''
ArtiV='120'
cursor.execute("INSERT INTO Clavesadd (Clave,Dato1,Usuario,usuFecha,usuHora,Dato2,Articulo,Unidad,Imagen) VALUES('"+ArtiV+"',0,'Sys','2022-02-07 00:00:00.000','13:00:00',0,0,0,0)")
'''

    

try:
    
    dfup=pd.read_excel('ARTICULOS ROTACION 2.xlsx')
    print('lectura de excel Ok')
    
except Exception as FaltaXLSX:
    print(FaltaXLSX)


#-----------------------------se importa la columna Articulo de clavesadd en un df
#query = "SELECT [CLAVE] FROM clavesadd;"
#dfArti = pd.read_sql(query, cnxn)

for j in range(0,1194):
    ArtiN=str(dfup.loc[j,'Nuevo'])
    #print(ArtiV)
    cursor.execute("INSERT INTO Clavesadd (Clave,Dato1,Usuario,usuFecha,usuHora,Dato2,Articulo,Unidad,Imagen) VALUES('"+ArtiN+"',0,'Sys','2022-02-07 00:00:00.000','13:00:00',0,0,0,0)")
    '''
    claveEx=0
    for i in range(0,1347):
        clave=str(dfArti.loc[i,'CLAVE'])
        #print(clave)
        if  clave==ArtiV:
            print('la clave ya existe', clave, ArtiV)
            claveEx=1
    if claveEx==0:
        print(ArtiV)
        cursor.execute("INSERT INTO Clavesadd (Clave,Dato1,Usuario,usuFecha,usuHora,Dato2,Articulo,Unidad,Imagen) VALUES('"+ArtiV+"',0,'Sys','2022-02-07 00:00:00.000','13:00:00',0,0,0,0)")
    #else:    
    #cursor.execute("INSERT INTO Clavesadd (Clave,Dato1,Usuario,usuFecha,usuHora,Dato2,Articulo,Unidad,Imagen) VALUES('"+ArtiV+'",0,Sys,2022-03-23 00:00:00.000,17:00:00,0,0,0,0)', row.Articulo)
    #cursor.execute("INSERT INTO Clavesadd (Clave,Dato1,Usuario,usuFecha,usuHora,Dato2,Articulo,Unidad,Imagen) VALUES('"+ArtiV+"',0,'Sys','2022-02-07 00:00:00.000','13:00:00',0,0,0,0)")
    #print('se crea clave', clave, ArtiV)
    '''
    
    
cnxn.commit()
cursor.close()
