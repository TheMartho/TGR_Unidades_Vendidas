from RPA.PDF import PDF
from robot.libraries.String import String
from RPA.Excel.Files import Files;
import re
import os
from datetime import date
from datetime import datetime
fecha_actual = datetime.now()
lib = Files()
pdf = PDF()
string = String()

def masterlibros():
    lib.open_workbook('Data\Base de Existencias Unidades.xlsx')       
    lib.read_worksheet("Export")       
    DtableFinal=lib.read_worksheet_as_table(name="Export",header=True, start=1).data
    return DtableFinal

def extract_data_from_first_page(contenido,filtro,carpeta):
    print("Buscando archivo : "+"PDF/"+carpeta+"/"+contenido)
    text = pdf.get_text_from_pdf("PDF/"+carpeta+"/"+contenido )
    pdf.close_pdf()
    
    if str(text).__contains__(filtro)  :
        print("TRUE")   
        print("Eliminando contenido de la carpeta ", contenido)
        os.remove("PDF/"+carpeta+"/"+contenido)

    else: 
     print("false")
    lugar= "PDF/"+carpeta+"/"+contenido
  
    año=fecha_actual.strftime('%Y')
  


    ubi="PDF/"+carpeta+"/"+contenido   
    ubi=str(ubi).replace("-2023","")
    ubi=ubi[0:42]
    os.rename(str("PDF/"+carpeta+"/"+contenido), str(ubi+" "+ año +".pdf") )
   

def filtroCuota(carpeta):
    
    f=open('Log Scraping/'+carpeta+'.txt',"r")
    fecha_actual = datetime.now()
    fecha_formateada = fecha_actual.strftime('%Y')
    Cuo=[]
    añocuo=[]
    a=0
    subtotal=0
    for a in f:
            if a.__contains__(fecha_formateada):
                Cuo.append(a[0:2].replace("-",""))
                añocuo.append(a[2:7].replace("-",""))
    mincuo=max(Cuo)
    añomax=max(añocuo)
    print("Filtra cuota a pagar")  
    filtrocuota=mincuo+"-"+añomax
    print(filtrocuota)
    return filtrocuota



def eliminarpdf(carpeta):
 filtro=filtroCuota(carpeta)
 print("filtro: "+ filtro)
 buscar=str(filtro[0:2])+str(filtro[4:6])
 print(buscar)
 contenido = os.listdir('PDF/'+carpeta+"/")
 print(contenido)
 for x in contenido:
    print(x)
    extract_data_from_first_page(x,buscar,carpeta)


def task():
    Dt=masterlibros()
    for dtable in Dt:

            Rut=str(dtable[5])
            nombre_consolidado=str(dtable[3])
            Inmobiliaria=dtable[1]
            Asset=dtable[2]
            region=dtable[1]
            comuna_=dtable[2]
            Carpeta=str(dtable[2]+" "+dtable[7]+" -"+dtable[8])
            Hoja=dtable[8]
            Activo=dtable[5]
            region=dtable[1]
            rolmatriz=dtable[7]
            rol1=dtable[7]                               
            rol2=dtable[8]
            Codigo=dtable[14]
            comuna=dtable[14]
            
                
           
            #try:eliminarpdf(Carpeta)
            #except FileNotFoundError: print("El archivo ,"+ Carpeta + ".txt → No fue contrado ")
            #except TypeError:print("no encontro nada .")
            #except UnboundLocalError:pass
        