import defRPAselenium
import moldesTerrenos
import modelsUnidadesvendidas
import modelsContribucionesTerrenos
import modelsSolicitudPago
import pandas as pd
from RPA.Browser.Selenium import Selenium;
import os
from shutil import rmtree
import time
from RPA.Desktop import Desktop
from datetime import datetime
import correo
import openpyxl

browser = Selenium()
try:
 moldesTerrenos.task_Modelos()
 moldesTerrenos.Asignaconsultafecha()
 #moldesTerrenos.task_Modelos()
 #moldesTerrenos.Asignaconsultafecha()
 pass
except:
    pass

Dt=moldesTerrenos.masterlibros()
urlbase=defRPAselenium.Pyasset(asset="base")
UrlMacro=defRPAselenium.Pyasset(asset="Ruta ")
libro=defRPAselenium.Pyasset(asset="LIBRO ")



def eliminarcarpetas():
    try:
        rmtree("PDF")
        rmtree("CSV")
        rmtree("Log Scraping")
        rmtree("Formato Solicitud")
        rmtree("Salida")
        os.remove("Excel\\Formato Reporte.xlsx")
        print("Eliminamos carpetas")

    except:
        pass

def Creacionescarpetas():
    print("Creado las carpetas para PDF's")

    try:
        os.mkdir('PDF')
        os.mkdir('CSV')
        os.mkdir("Formato Solicitud")
        os.mkdir("Log Scraping")
        os.mkdir("Salida") 
        os.remove("Roles No encontrados.txt")
    except:
        pass




def repetidos(carpeta):
    with open('Repetidos.txt', 'a') as archivo:
            # Escribir el número en el archivo, seguido de un salto de línea
            archivo.write(str(carpeta)+'\n')
    # Informar que se han escrito los números en el archivo
    print("Se agrego el Rol Repetido")


def taskRut():
        agregar=defRPAselenium.Pyasset(asset="Agregar")
        nombreActual=""
        for dtable in Dt:
            if dtable[15] == "SI":
                Carpeta=None
                if(len(str(dtable[8])))==1:
                    rolCompleto=str(str(dtable[7])+"-00"+str(dtable[8]))
                if(len(str(dtable[8])))==2:
                    rolCompleto=str(str(dtable[7])+"-0"+str(dtable[8]))
                if(len(str(dtable[8])))==3:
                    rolCompleto=str(str(dtable[7])+"-"+str(dtable[8]))
                Hoja=dtable[8]
                Carpeta=str(dtable[2]+" "+str(dtable[7])+"-"+str(dtable[8]))
                region=dtable[1]
                rol1=dtable[7]                               
                rol2=dtable[8]
                unidad=dtable[11]
                comuna=dtable[14]   
                nombreConsolidado=str(dtable[3])
                consulta=True
                nIntentos=1
                defRPAselenium.LOGconsulta(region,comuna,rol1,rol2)
                if os.path.exists("Log Scraping/"+Carpeta+".txt") and str(agregar)=="True":
                    print("Estos datos ya se consultaron")
                else:
                    rut=modelsSolicitudPago.getRut(nombreConsolidado)
                    while consulta==True:
                        try:
                            print("Realizando Intento N°"+str(nIntentos))
                            if nombreConsolidado != nombreActual:
                                    try:
                                        driver.switch_to.default_content()
                                        driver.quit()
                                    except:pass
                                    nombreActual=nombreConsolidado
                                    valores=modelsSolicitudPago.getCredentials(nombreConsolidado)
                                    user, nomInmobiliaria = valores
                                    cla=modelsSolicitudPago.getPass(nomInmobiliaria)
                                    try:defRPAselenium.creacioncarpetas(nomInmobiliaria)
                                    except:pass
                                    try:
                                        defRPAselenium.valoresReporte(rut,nomInmobiliaria,str(unidad))
                                        driver=defRPAselenium.abrirNavegadorRut(user,cla)
                                        defRPAselenium.busquedaRol(driver,Carpeta,rol1,rol2,rolCompleto,nomInmobiliaria)
                                        
                                    except:
                                        driver.switch_to.default_content()
                                        try:driver.quit()
                                        except:pass
                                        raise ValueError("Error")
                            else:
                                inmobi=modelsSolicitudPago.getInmobiliaria(nombreConsolidado)
                                defRPAselenium.valoresReporte(rut,inmobi,str(unidad))
                                defRPAselenium.busquedaRol(driver,Carpeta,rol1,rol2,rolCompleto,nomInmobiliaria,str(unidad))
                            consulta=False
                        except:
                            nombreActual=""
                            try:
                                driver.switch_to.default_content()
                                driver.quit()
                            except:pass
                            r_completo=str(rol1)+"-"+str(rol2)
                            defRPAselenium.reportarError("Fallo al navegar por el menú de TGR, Contribuciones por RUT - Reintento N°"+str(nIntentos-1),True,rolCompleto,unidad)
                            consulta=True
                            nIntentos+=1
        try:
            driver.switch_to.default_content()
            driver.quit()
        except:pass


                     



def buscarPorRol(rol1,rol2):
    for dtable in Dt:
            rolMatriz=str(dtable[7]).strip()                          
            rolUnidad=str(dtable[8]).strip()
            if rol1 == rolMatriz and rol2 == rolUnidad:
                comuna=dtable[14]
                region=dtable[1]
                Carpeta=str(dtable[2]+" "+str(dtable[7])+"-"+str(dtable[8]))
                nombreConsolidado=str(dtable[3])
                unidad = dtable[11]
                return comuna, region, Carpeta, nombreConsolidado, unidad



def taskRol():
    # Abre el archivo en modo lectura
    agregar=defRPAselenium.Pyasset(asset="Agregar")
    with open('Roles No encontrados.txt', 'r') as archivo:
        # Itera sobre cada línea del archivo
        for linea in archivo:
            # Procesa cada línea 
            partes=linea.strip()
            if len(partes) >= 2:
                # Asignar las partes a variables
                roles=partes.split()
                rol1=roles[0]
                rol2=roles[1]
                variables=buscarPorRol(rol1,rol2)
                comuna, region, Carpeta, nombreConsolidado, unidad = variables
                nomInmobiliaria=modelsSolicitudPago.getInmobiliaria(nombreConsolidado)
                rut=modelsSolicitudPago.getRut(nombreConsolidado)
                consulta=True
                nIntentos=1
                if os.path.exists("Log Scraping/"+Carpeta+".txt") and str(agregar)=="True":
                    print("Archivo ya consultado")
                    consulta=False
                while consulta==True:
                    defRPAselenium.LOGconsulta(region,comuna,rol1,rol2)
                    try:
                        defRPAselenium.valoresReporte(rut,nomInmobiliaria,str(unidad))
                        print("Realizando Intento N°"+str(nIntentos))
                        tabla = defRPAselenium.navegacion(str(region),str(comuna),str(rol1),str(rol2),str(Carpeta),str(nomInmobiliaria),str(unidad))
                        consulta=False
                        r_completo=str(rol1)+"-"+str(rol2)
                    except:
                        consulta=True
                        try:tabla.quit()
                        except:pass
                        r_completo=str(rol1)+"-"+str(rol2)
                        defRPAselenium.reportarError("Fallo al navegar por el menú de TGR, Contribuciones por ROL - Reintento N°"+str(nIntentos-1),True,r_completo,unidad)
                        nIntentos+=1


"""def repetidosOrden():
    with open('Roles No encontrados.txt', 'r') as archivo:
        # Itera sobre cada línea del archivo
        woorkbook=openpyxl.Workbook()
        sheet=woorkbook.active
        fila=['Carpeta', 'Nombre Consolidados', 'Comuna', 'Region', 'Rol']
        sheet.append(fila)
        for linea in archivo:
            # Procesa cada línea 
            partes=linea.strip()
            if len(partes) >= 2:
                # Asignar las partes a variables
                roles=partes.split()
                rol1=roles[0]
                rol2=roles[1]
                variables=buscarPorRol(rol1,rol2)
                comuna, region, Carpeta, nombreConsolidado = variables
                fila=[str(Carpeta),str(nombreConsolidado),comuna,region,str(roles)]
                sheet.append(fila)
        woorkbook.save('Excel Repetidos.xlsx')"""




def tgc():
 taskRut()
 try:
    taskRol()
 except:pass



def enviandoCorreo(fecha_formateada_inicio, fecha_formateada_final):
    archivo_excel = "Excel/Formato Reporte.xlsx"
    nombre_hoja = 'Errores'
    df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)
    if not df.empty:
        print("Enviando Correo con errores")
        correo.enviarCorreo(fecha_formateada_inicio, fecha_formateada_final,True)
    else:
        print("Enviando Correo sin errores")
        libro_excel=openpyxl.load_workbook(archivo_excel)
        hoja_a_borrar=libro_excel["Errores"]
        libro_excel.remove(hoja_a_borrar)
        libro_excel.save(archivo_excel)
        correo.enviarCorreo(fecha_formateada_inicio, fecha_formateada_final,False)
    


def borrarPDF():
        pdfs="D:/.PROGRAMACION/Timix/El Bueno/tgr_Unidades_vendidas/PDF"
        inmobiliarias=os.listdir(pdfs)
        for inmo in inmobiliarias:
            print(inmo)
            ruta="D:/.PROGRAMACION/Timix/El Bueno/tgr_Unidades_vendidas/PDF/"+inmo
            archivos = os.listdir(ruta)
            for archivo in archivos:
                    existe=False
                    partes=str(archivo).split()
                    try:
                        rol=partes[3]
                    except:
                        rol="999-999"
                    for dtable in Dt:
                        rolCompleto=str(str(dtable[7])+"-"+str(dtable[8]))
                        if rol==rolCompleto:
                            existe=True
                    if existe==False:
                        print("Archivo no deberia existir: "+archivo)
                        os.remove("D:/.PROGRAMACION/Timix/El Bueno/tgr_Unidades_vendidas/PDF/"+inmo+"/"+archivo)


def procesoCompletado():
    import tkinter as tk
    from tkinter import messagebox

    # Crear una ventana principal
    root = tk.Tk()
    root.title("Ejemplo Tkinter")

    # Establecer el tamaño de la ventana
    root.geometry("300x200")

    # Establecer la ubicación de la ventana en la pantalla (posiciónX, posiciónY)
    root.geometry("+500+300")  # Ajusta estas coordenadas según tus necesidades
    # Hacer que la ventana esté siempre en la parte superior
    root.attributes("-topmost", True)

    # Función para mostrar el cuadro de mensaje
    def mostrar_mensaje():
        messagebox.showinfo(message="Ejecución Finalizada", title="Ejecución Bot Unidades")

    # Botón para mostrar el mensaje
    boton_mostrar = tk.Button(root, text="Mostrar Mensaje", command=mostrar_mensaje)
    boton_mostrar.pack(pady=20)

    # Iniciar el bucle principal de Tkinter
    root.mainloop()


if __name__ == "__main__":
   tiempoInicio=time.time()
   fecha_inicio=datetime.now()
   #fecha_formateada_inicio = fecha_inicio.strftime('%d/%m/%Y %H:%M:%S')
   fecha_formateada_inicio="06/12/2023 11:31:30"
   eliminarcarpetas()
   Creacionescarpetas()
   defRPAselenium.bakup()
   tgc()
   modelsUnidadesvendidas.task()
   modelsSolicitudPago.task()
   defRPAselenium.salida()
   try:borrarPDF()
   except:pass
   tiempoFinal=time.time()
   TiempoTotal=tiempoFinal-tiempoInicio
   fecha_final=datetime.now()
   fecha_formateada_final = fecha_final.strftime('%d/%m/%Y %H:%M:%S')
   #enviandoCorreo(fecha_formateada_inicio, fecha_formateada_final)
   print("Tiempo total de ejecucion es " + str(TiempoTotal) + " seg")
   procesoCompletado()
