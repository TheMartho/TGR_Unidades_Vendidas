from RPA.Tables import Tables
import modelsSolicitudPago
import modelsContribucionesTerrenos
import defRPAselenium
from RPA.Excel.Files import Files;
import pandas as pd
import re
import os
import warnings
lib = Files()
library = Tables()
from datetime import datetime
fecha_actual = datetime.now()


def ResumenLibro(inmobiliaria):
    try:
        lib.open_workbook("Salida/resumen unidades vendidas.xlsx")      #ubicacion del libro
        lib.read_worksheet(inmobiliaria)                                              #nombre de la hoja
        dtresumen=lib.read_worksheet_as_table(name=inmobiliaria,header=True, start=1).data
        return dtresumen
    except:
        return None

def excelExistenciaUnidades():
    warnings.simplefilter(action='ignore', category=UserWarning)
    excel_file = pd.ExcelFile('Data/Base de Existencias Unidades.xlsx')
    # Lee la hoja como un DataFrame
    df = excel_file.parse("Export")
    # Convierte el DataFrame en una lista de diccionarios
    DtableFinal = df.to_dict(orient='records')
    return DtableFinal

def masterlibros():
    lib.open_workbook('Data\Base de Existencias Unidades.xlsx')       
    lib.read_worksheet("Export")       
    DtableFinal=lib.read_worksheet_as_table(name="Export",header=True, start=1).data
    return DtableFinal

def Consolidado(txt,nombre_consolidado,Rut_Cliente,Region,comuna,ROMATRIZ,Rol_Unidad,tipoProducto,numero):

    ubicacion="Salida/resumen unidades vendidas.xlsx"
    nomInmobiliaria=modelsSolicitudPago.getInmobiliaria(nombre_consolidado)
    nombre_consolidado_HOJA=nomInmobiliaria[0:31]


    f=open("Log Scraping/"+txt+".txt","r")
    file_content=f.read()

    lines = file_content.split('\n')
    normalized_lines = [' '.join(re.sub(r'\s+', ' ', line).split()) for line in lines]
    normalized_content = '\n'.join(normalized_lines)
    
    # Dividir el contenido del archivo en líneas
    lines = normalized_content.split('\n')

    # Eliminar líneas en blanco al principio y al final
    lines = [line.strip() for line in lines if line.strip()]

    # Unir las líneas nuevamente en un solo string
    normalized_content = '\n'.join(lines)
    f.close()
    with open("CSV/"+txt+".csv","w") as w:
        w.write(normalized_content)




    orders = library.read_table_from_csv(
        "CSV/"+txt+".csv",header=False,delimiters=" ")


    CapturaSCRAPIADO= ([{ }])
        
    #mes_string=int(fecha_actual.strftime('%m'))
    #mes_string=int(mes_string-1)
    
    filtroCuota=defRPAselenium.filtrarCuota()
    intFiltro=int(filtroCuota[0:1])
    intAnioFiltro=int(filtroCuota[2:])
    try:
        for x in orders:
            if len(x[0]) != 0 and str(x[0]!="No se encontraron Deudas"):
                print(comuna+" "+ROMATRIZ+"-"+Rol_Unidad)
                CUOTA=x[0]
                VALOR_CUOTA=x[1]
                TOTAL_A_PAGAR=str(x[4]).replace(".","")
                try:
                    strCuota=str(CUOTA)
                    cutCuota=strCuota[0:1]
                    intCuota=int(cutCuota)
                    anioCuota=int(strCuota[2:])
                except:
                    intCuota=99
                if intCuota<=intFiltro or anioCuota<intAnioFiltro:
                    CapturaSCRAPIADO.append({
                        'INMOBILIARIA':nomInmobiliaria,
                        'REGION':Region,
                        'COMUNA':comuna,
                        'ROL MATRIZ':ROMATRIZ,
                        'ROL UNIDAD':Rol_Unidad,
                        'TIPO PRODUCTO':tipoProducto,
                        'NUMERO':numero,
                        'CUOTA':CUOTA,
                        'TOTAL A PAGAR':"$"+(TOTAL_A_PAGAR)    
                            })
    except:
        print("Ocurrio un error con "+comuna+" "+ROMATRIZ+"-"+Rol_Unidad)
    try:lib.open_workbook(ubicacion) 
    except TypeError: lib.create_workbook(ubicacion) 
    except FileNotFoundError: lib.create_workbook(ubicacion) 

    if lib.worksheet_exists(nombre_consolidado_HOJA)==True:
        print("existe")
    else:
        print("no existe crea uno nuevo")
        lib.create_worksheet(nombre_consolidado_HOJA)

    #variables de fechas 
    #año_string=fecha_actual.strftime('%Y')
    #mes_string=fecha_actual.strftime('%m')


    #encabezado
    lib.set_cell_value(1,1,"INMOBILIARIA")
    lib.set_cell_value(1,2,"REGION")
    lib.set_cell_value(1,3,"COMUNA")
    lib.set_cell_value(1,4,"ROL MATRIZ")
    lib.set_cell_value(1,5,"ROL UNIDAD")
    lib.set_cell_value(1,6,"TIPO PRODUCTO")
    lib.set_cell_value(1,7,"NUMERO")
    lib.set_cell_value(1,8,"CUOTA")
    lib.set_cell_value(1,9,"TOTAL A PAGAR")

    # Introduccimos los datos con append  
    lib.set_active_worksheet(nombre_consolidado_HOJA)
    lib.append_rows_to_worksheet(content=CapturaSCRAPIADO,header=False,start=1)

    # Limpiamos la data
    lib.set_active_worksheet(nombre_consolidado_HOJA)
    lib.read_worksheet(nombre_consolidado_HOJA)       #nombre de la hoja
    lista=lib.read_worksheet_as_table(name=nombre_consolidado_HOJA,header=True, start=1).data
    registros=(lib.find_empty_row()*10)




    # eliminamos los espacios vacios en las celdas
    for celdas in range(registros) :

        Buscar_vacias =lib.get_cell_value(1+celdas,"A")
        Buscar_Totales =lib.get_cell_value(1+celdas,"D")
        buscar_fechas = lib.get_cell_value(1+celdas,"D")
       

        if Buscar_vacias is None or Buscar_Totales=="total":
            lib.delete_rows(celdas+1)
        

    # hacemos un segundo barrido
    for celdas in range(registros):
        Buscar_vacias =lib.get_cell_value(1+celdas,"A")
        if Buscar_vacias is None or Buscar_Totales=="total":
            lib.delete_rows(celdas+1)



     # hacemos un tercer barrido
    for celdas in range(registros):
        Buscar_vacias =lib.get_cell_value(1+celdas,"A")
        if Buscar_vacias is None or Buscar_Totales=="total":
            lib.delete_rows(celdas+1)
    
    lib.save_workbook(ubicacion)


            

                
def task():
    Dt=masterlibros()
    for dtable in Dt:
        if dtable[15] == "SI":
            Rut=str(dtable[5])
            nombre_consolidado=str(dtable[3])
            Inmobiliaria=dtable[1]
            region=dtable[1]
            comuna_=dtable[2]
            Carpeta=str(dtable[2]+" "+str(dtable[7])+"-"+str(dtable[8]))
            region=dtable[1]
            rol1=dtable[7]                               
            rol2=dtable[8]
            tipoProducto=dtable[10]
            numero=dtable[11]

            try:Consolidado(Carpeta,nombre_consolidado,Rut,region,comuna_,rol1,rol2,tipoProducto,numero)
            except FileNotFoundError: print("El archivo ,"+ Carpeta + ".txt → No fue contrado ")
            except TypeError:print("no encontro nada .")
            except UnboundLocalError:pass

    nombreInmobiliariaActual=""
    u=excelExistenciaUnidades()
    for fila in u:
        nombreConsolidado=str(fila['nombre_consolidado'])
        Inmobiliaria=modelsSolicitudPago.getInmobiliaria(nombreConsolidado)
        Inmobiliaria=modelsSolicitudPago.getInmobiliaria(nombreConsolidado)
        if Inmobiliaria != nombreInmobiliariaActual:
            nombreInmobiliariaActual=Inmobiliaria
            totales(Inmobiliaria,nombreConsolidado)







def totales(inmobiliaria,nombreConsolidado):
    totalDefinitivo=0
    contador=0
    fila=0
    subtotal=0
    resumen=ResumenLibro(inmobiliaria[0:31])
    if resumen is None:
        print("No hay Valores para calcular totales")
    else:
        tamaño=len(resumen)*2
        #filasNone=0
        filaAntesNone=0
        while fila <= tamaño:
            fila+=1
            print(str(fila))
            if str(lib.get_cell_value(fila,"I"))=='TOTAL A PAGAR' or str(lib.get_cell_value(fila,"I"))=="" or lib.get_cell_value(fila,"I")==None:
                print(lib.get_cell_value(fila,"I"))
            else:
                if str(lib.get_cell_value(fila,"H"))!='TOTAL':
                    pagar=str(lib.get_cell_value(fila,"I"))
                    subtotal=subtotal+int(pagar[1:])
                    totalDefinitivo=totalDefinitivo+int(pagar[1:])
                    contador+=1
                    filaAntesNone=fila


            if contador==10:
                lib.insert_rows_after(fila,2)
                lib.set_cell_value(fila+1,"H","TOTAL")
                lib.set_cell_value(fila+1,"I",subtotal,fmt="[$$-es-CL]#,##0")
                lib.save_workbook()
                subtotal=0
                contador=0


        fila=0
        total=0
        esTotal=lib.get_cell_value(filaAntesNone+1,"H")
        esTotal=str(esTotal)
        if esTotal!="TOTAL":
            while fila!=10:
                if lib.get_cell_value(filaAntesNone-fila,"I") is None:
                    break
                else:
                    try:
                        pagar=lib.get_cell_value(filaAntesNone-fila,"I")
                        total=total+int(pagar[1:])
                        fila+=1
                    except:
                        break

            lib.insert_rows_after(filaAntesNone+2,2)
            lib.set_cell_value(filaAntesNone+2,"H","TOTAL")
            lib.set_cell_value(filaAntesNone+2,"I",total,fmt="[$$-es-CL]#,##0")        
            lib.save_workbook()
            lib.close_workbook()
        modelsContribucionesTerrenos.task(inmobiliaria,totalDefinitivo,nombreConsolidado)

                               

            