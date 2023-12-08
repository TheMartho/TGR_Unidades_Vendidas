import modelsUnidadesvendidas
import moldesTerrenos
import defRPAselenium
import pandas as pd
import warnings

def getPass(nomInmobiliaria):
    m=moldesTerrenos.excelMaster()
    cla=""
    for fila in m:
        if str(fila['Inmobiliaria']) == str(nomInmobiliaria):
            cla=str(fila['A'])
            break
    return cla

def getH(consolidado):
    #Obtenemos contribuciones
    m=moldesTerrenos.excelInmobiliarias()
    numero=0
    for fila in m:
        if str(fila['Nombre Consolidado']) == consolidado:
            numero=(fila['N° '])
            break
    h= int(float(numero))
    return h

def getCredentials(consolidado):
    #Obtenemos contribuciones
    m=moldesTerrenos.excelInmobiliarias()
    user=""
    nomInmobiliaria=""
    for fila in m:
        if str(fila['Nombre Consolidado']) == str(consolidado):
            user=str(fila['RUT'])
            nomInmobiliaria=str(fila['Inmobiliaria'])
            break
    return user,nomInmobiliaria

def getInmobiliaria(consolidado):
    #Obtenemos contribuciones
    m=moldesTerrenos.excelInmobiliarias()
    nomInmobiliaria=""
    for fila in m:
        if str(fila['Nombre Consolidado']) == str(consolidado):
            nomInmobiliaria=str(fila['Inmobiliaria'])
            break
    return nomInmobiliaria

def getRut(consolidado):
    #Obtenemos contribuciones
    m=moldesTerrenos.excelInmobiliarias()
    rut=""
    for fila in m:
        if str(fila['Nombre Consolidado']) == str(consolidado):
            rut=str(fila['RUT'])
            break
    return rut

def excelUnidadesVendidas(nombre_consolidado):
    warnings.simplefilter(action='ignore', category=UserWarning)
    try:
        excel_file = pd.ExcelFile('Salida/resumen unidades vendidas.xlsx')
        # Lee la hoja como un DataFrame
        df = excel_file.parse(nombre_consolidado[0:31])
        # Convierte el DataFrame en una lista de diccionarios
        DtableFinal = df.to_dict(orient='records')
        return DtableFinal
    except:
        return None
    

def makeCupon(nombre_consolidado):
    numeroSolicitud=1
    inmobiliaria=getInmobiliaria(nombre_consolidado)
    c=excelUnidadesVendidas(inmobiliaria)
    if c is None:
     print("---------------------------------------------")
    else:
        h=getH(nombre_consolidado)
        for fila in c:
            if str(fila['CUOTA']) == nombre_consolidado or str(fila['CUOTA']) == 'TOTAL' or str(fila['CUOTA']) == 'CUOTA':
                    if str(fila['CUOTA']) == 'TOTAL':
                        total=str(fila['TOTAL A PAGAR'])
                        defRPAselenium.formatosolicitusd(h,total,numeroSolicitud)
                        print("Se creo la Solicitud de Pago N°"+str(numeroSolicitud))
                        numeroSolicitud=numeroSolicitud+1
            else:
                    print("__________________________________________")




def task():
    nombreActual=""
    Dt = modelsUnidadesvendidas.masterlibros()
    for dtable in Dt:
        nombre_consolidado=str(dtable[3])
        if nombre_consolidado!= nombreActual:
            print("Creando Solicitudes de Pago de: "+ nombre_consolidado)
            makeCupon(str(nombre_consolidado))
            nombreActual=nombre_consolidado  
