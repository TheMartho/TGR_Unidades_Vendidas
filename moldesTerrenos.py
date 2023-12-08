from RPA.Browser.Selenium import Selenium;
from RPA.Excel.Application import Application
from RPA.Windows import Windows
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files;
import time
import warnings
import random
import os
import shutil
from datetime import date
from datetime import datetime, timedelta
from RPA.Tables import Tables
import pandas as pd
library = Tables()
lib = Files()
fecha_actual = datetime.now()

def masterlibros():
    lib.open_workbook('Data\Base de Existencias Unidades.xlsx')       
    lib.read_worksheet("Export")       
    DtableFinal=lib.read_worksheet_as_table(name="Export",header=True, start=1).data
    return DtableFinal

def excelResumen():
    warnings.simplefilter(action='ignore', category=UserWarning)
    # Abre el archivo de Excel
    excel_file = pd.ExcelFile('Salida/resumen unidades vendidas.xlsx')
    # Inicializa un diccionario para almacenar los datos de todas las hojas
    datos_por_hoja = {}
    
    # Itera a través de todas las hojas en el archivo Excel
    for hoja_nombre in excel_file.sheet_names:
        # Lee cada hoja como un DataFrame
        df = excel_file.parse(hoja_nombre)
        
        # Convierte el DataFrame en una lista de diccionarios
        datos_por_hoja[hoja_nombre] = df.to_dict(orient='records')

    return datos_por_hoja

def excelMaster():
    warnings.simplefilter(action='ignore', category=UserWarning)
    excel_file = pd.ExcelFile('Data/Master.xlsx')
    # Lee la hoja como un DataFrame
    df = excel_file.parse("Listado")
    # Convierte el DataFrame en una lista de diccionarios
    DtableFinal = df.to_dict(orient='records')
    return DtableFinal

def excelInmobiliarias():
    warnings.simplefilter(action='ignore', category=UserWarning)
    excel_file = pd.ExcelFile('Data/Inmobiliarias.xlsx')
    # Lee la hoja como un DataFrame
    df = excel_file.parse("Resumen")
    # Convierte el DataFrame en una lista de diccionarios
    DtableFinal = df.to_dict(orient='records')
    return DtableFinal


def Asignaconsultafecha():
    
    lib.open_workbook('Data\Base de Existencias Unidades.xlsx')       
    lib.read_worksheet("Export")       
    master=lib.read_worksheet_as_table(name="Export",header=True, start=1).data
    fecha_actual= datetime.now()
    anio_actual=fecha_actual.year
    primer_dia_mes_anterior=fecha_actual.replace(day=1) - timedelta(days=1)
    mes_anterior=primer_dia_mes_anterior.month
    lib.set_cell_value(1,"P","Consultar")

    celda=1
    try:
        for base in master:
            celda+=1
            escritura_correcta=False
            cadena_fecha_escritura=str(base[12])
            cadena_fecha_entrega=str(base[13])
            try:
                fecha_escritura=datetime.strptime(cadena_fecha_escritura, "%Y-%m-%d %H:%M:%S")
                fecha_escritura=fecha_escritura.strftime("%Y-%m-%d")
                fecha_escritura=datetime.strptime(fecha_escritura, "%Y-%m-%d")
                if fecha_escritura.month<=mes_anterior and fecha_escritura.year == anio_actual:
                    escritura_correcta=True
            except:
                escritura_correcta=True
            if escritura_correcta==True:
                try:
                    fecha_entrega=datetime.strptime(cadena_fecha_entrega, "%Y-%m-%d %H:%M:%S")
                    fecha_entrega=fecha_entrega.strftime("%Y-%m-%d")
                    fecha_entrega=datetime.strptime(fecha_entrega, "%Y-%m-%d")
                except:
                    try:
                        rol_matriz=str(base[7])
                        rol_unidad=str(base[8])
                        rol=int(rol_matriz)
                        rol=int(rol_unidad)
                        lib.set_cell_value(int(celda),"P","SI")
                    except:
                        rol=0
                        
    except:
            print("Ocurrio un error durante la asignacion de SI")
        
    lib.save_workbook()
    lib.close_workbook()
    print("Filtros Aplicados Correctamente")

def AsignaCodigoComuna():
    lib = Files()
    lib.open_workbook('Data\Base de Existencias Unidades.xlsx')       
    lib.read_worksheet("Export")       
    master=lib.read_worksheet_as_table(name="Export",header=True, start=1).data

    lib.set_cell_value(1,"O","Codigo comuna")

    celda=1
    
    for s in master:

        try:
            if str(s[8]) == "None":
             break
            else:
                celda=int(celda+1)
                comuna=str(s[2])
                codigo=CodigoComuna(comuna)
                lib.set_cell_value(int(celda),"O",codigo)

        except:
            pass
    lib.save_workbook()
    lib.save_workbook()
    lib.close_workbook()

def mesconsultar():
 fecha_actual = datetime.now()
 #fecha_formateada = fecha_actual.strftime('%d/%m/%Y')
 fecha_formateada = fecha_actual.strftime('%Y-%m')
 print(fecha_formateada)

def CodigoComuna(comuna):
    file_path = 'Data/Codigos Comunas.xlsx'
    #Lee el archivo excel
    df = pd.read_excel(file_path, sheet_name='codigos')
    # Filtra el DataFrame para encontrar la fila con la coincidencia de comuna
    consulta = df.loc[df['codigolibro'] == comuna, 'Out_comuna'].iloc[0]
    #Devuelve la consulta
    return consulta



def task_Modelos():
        dt=masterlibros()
        lib.set_cell_value(1,"O","Codigo comuna")
        fila=1
        for resumen in dt:
            fila+=1
            region=str(resumen[1])
            comuna=str(resumen[2])
            RolMatriz=str(resumen[7])
            print("Consultando "+ RolMatriz +" "+region+" - "+comuna)
            #AsignaCodigoComuna()
            codigo=CodigoComuna(comuna)
            lib.set_cell_value(fila,"O",str(codigo))
        lib.save_workbook()
        lib.close_workbook()
        print("Codigos de Comuna Asignados")

def logscraping(carpeta,Rolmatriz):
    f=open('Log Scraping/'+carpeta+".txt","r")
    CapturaSCRAPIADO= ([{ }])
    f=f.readlines()

    for x in f:
         CUOTA=x[0:7]
         VALOR_CUOTA=x[8:15]
         NRO_FOLIO=x[16:25]
         VENCIMIENTO=x[26:36]
         TOTAL_A_PAGAR=x[37:48]
         EMAIL=x[48:55]

        
         if str(x[48:55]) in " ":
            print("------------------------")
         else: 
             CapturaSCRAPIADO.append({
               'CUOTA':CUOTA,
               'VALOR CUOTA':int(VALOR_CUOTA), 
               'NRO FOLIO':NRO_FOLIO,
               'VENCIMIENTO':VENCIMIENTO,
               'TOTAL A PAGAR':TOTAL_A_PAGAR,
               'EMAIL':EMAIL
                 })
                
             
    lib.create_workbook() 
    lib.create_worksheet(Rolmatriz)
    lib.append_rows_to_worksheet(CapturaSCRAPIADO, header=True)
    lib.save_workbook('Excel/'+carpeta+".xlsx")
    lib.close_workbook()

def salida(carpeta,Rolmatriz,rut,inmobiliaria,Región,Comuna):

    """los datos necesarios son :
    carpeta: str
    Rolmatriz: str
    rut: str
    inmobiliaria: str
    Región: str
    Comuna: str

    """
   
    lib.open_workbook("Excel/"+carpeta+'.xlsx')        #ubicacion del libro
    lib.read_worksheet(str(Rolmatriz))       #nombre de la hoja
    outlista=lib.read_worksheet_as_table(name=str(Rolmatriz),header=True, start=1).data
    lib.close_workbook()
    recol=0
#encabezados
    tabla=([{ }])
    
    for x in outlista:

     if str(x[1])=="None":
        total=0

     else:

        try:
            tabla.append({
               'RUT':rut,
               'INMOBILIARIA':inmobiliaria, 
               'Región':Región,
               'cuota':str(x[0]),
               'Comuna':Comuna,
               'Rolmatriz':Rolmatriz,
               'Informacion Tesoreria':"Informacion_Tesoreria",
               'Monto':str(x[1])
                
                 })
         
            total=total+int(x[1])
            print(total)

        except:
            pass
#Diligenciamos los totales        
    tabla.append({

               'RUT':"",
               'INMOBILIARIA':"", 
               'Región':"",
               'cuota':"",
               'Comuna':"",
               'Rolmatriz':"",
               'Informacion Tesoreria':"total",
               'Monto':str(total)
                
                 })
        
    try:   
            #open("Salida\SalidaUnidadesVendidas.xlsx")
            lib.open_workbook("Salida\SalidaUnidadesVendidas.xlsx")
            print("el libro existe")
            Existe=lib.worksheet_exists(rut)
            print(Existe)
            
            if Existe == True:
                 
                lib.read_worksheet("Salida\SalidaUnidadesVendidas.xlsx")
                DtableFinal=lib.read_worksheet_as_table(name=rut,header=True, start=1).data
                lib.append_rows_to_worksheet(tabla, header=True)
    except:
            print("el libro no existe")
            Existe=False
            lib.create_workbook("Salida\SalidaUnidadesVendidas.xlsx")
            lib.create_worksheet(rut)
            lib.append_rows_to_worksheet(tabla, header=True)
            
        

       
    lib.save_workbook()
    lib.close_workbook()
    tabla=([{ }]) 


"""
carpeta="Nunoa 5701-477"
Rolmatriz="4505-54"
rut="11111111"
inmobiliaria="ejemplo"
Región="metropolitana "
Comuna="nunoa"


f=open('Log Scraping/'+carpeta+".txt","r")
f=f

"""






