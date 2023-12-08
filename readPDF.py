
import os
import PyPDF2
import re

carpeta="Excel/"

def convertir():
    ruta= os.path.join(carpeta, 'nuevo.txt')
    # Abre el archivo PDF en modo de lectura binaria
    with open('PDF\Santiago 1214 -14\Cupon de pago 1214-14 4-2023.pdf', 'rb') as pdf_file:
        # Crea un objeto PDFReader
        pdf_reader = PyPDF2.PdfReader(pdf_file)

        # Inicializa una cadena vacía para almacenar el texto
        text = ''

        # Itera a través de las páginas del PDF
        for page_num in range(len(pdf_reader.pages)):
            page = pdf_reader.pages[page_num]
            text += page.extract_text()

    # Guarda el texto en un archivo TXT
    with open('Excel/archivo.txt', 'w', encoding='utf-8') as txt_file:
        txt_file.write(text)


def encontrarTexto():

    # Abre el archivo de texto en modo de lectura
    with open('Excel/archivo.txt', 'r', encoding='utf-8') as txt_file:
        # Lee el contenido del archivo en una variable
        text = txt_file.read()

    # Define el patrón que deseas buscar
    patron = r'001001214-.*'  

    # Busca el patrón en el texto
    resultados = re.findall(patron, text)

    # Imprime los resultados
    if resultados:
        print("Se encontraron los siguientes patrones:")
        for resultado in resultados:
            Rol=resultado[10:13]
            print(Rol)
            print(resultado)
            
            
    else:
        print("No se encontraron patrones que coincidan.")
