#import fitz  # PyMuPDF
import pandas as pd
import os

# Función para extraer texto del PDF
"""
def extraer_texto_pdf(ruta_pdf):
    doc = fitz.open(ruta_pdf)
    texto = ""
    for pagina in doc:
        texto += pagina.get_text()
    doc.close()
    #print(texto)
    return texto
"""
import pdfplumber

def extraer_texto_pdf(ruta_pdf):
    texto = ""
    with pdfplumber.open(ruta_pdf) as pdf:
        for pagina in pdf.pages:
            texto += pagina.extract_text() or ""
        print(texto)
    return texto

"""
def procesar_texto(texto):
    lineas = texto.split('\n')
    nombre_escuela, grado, grupo = "", "", ""
    alumnos = []  # Almacenará tuplas de (CURP, Nombre)
    captura_datos_alumno = False

    i = 0  # Inicializar el contador para el bucle while
    while i < len(lineas):
        linea = lineas[i]
        if "ESCUELA:" in linea:
            nombre_escuela = lineas[i+2].strip()
        elif "GRADO:" in linea:
            grado = lineas[i+11].strip()
        elif "GRUPO:" in linea:
            grupo = lineas[i+11].strip()
        elif "FIRMA DIRECTOR" in linea:
            break
        elif "LISTA DE ALUMNOS A LA FECHA" in linea:
            captura_datos_alumno = True
            i += 2  # Saltar "MATRÍCULA" y el primer número de matrícula
            continue
        elif captura_datos_alumno:
            if i + 2 < len(lineas):  # Asegurar que hay suficientes líneas restantes para CURP y nombre
                curp = lineas[i+1].strip()
                nombre = lineas[i+2].strip()
                alumnos.append((curp, nombre))
                i += 4  # Saltar al siguiente bloque de alumno
                continue
        i += 1  # Incrementar el contador para continuar con la siguiente línea

    return nombre_escuela, grado, grupo, alumnos
"""

def procesar_texto(texto):
    lineas = texto.split('\n')
    nombre_escuela, grado, grupo = "", "", ""
    alumnos = []  # Almacenará tuplas de (CURP, Nombre)

    # Extraer información de la escuela, grado y grupo
    for linea in lineas:
        if "ESCUELA:" in linea:
            nombre_escuela = linea.split("ESCUELA:")[1].split("GRUPO:")[0].strip()
            grupo = linea.split("GRUPO:")[1].strip()
        if "GRADO:" in linea:
            grado = linea.split("GRADO:")[1].strip()

    # Iniciar la extracción de alumnos después de 'NOMBRE DEL ALUMNO'
    inicio = False
    for linea in lineas:
        if "NOMBRE DEL ALUMNO" in linea:
            inicio = True
            continue
        if inicio:
            if linea.strip() == "FIRMA DIRECTOR NOMBRE Y FIRMA MAESTRO":
                break
            partes = linea.split()
            if len(partes) > 1:  # Verificar que la línea contiene más de un elemento
                curp = partes[2]  # CURP es el tercer elemento
                nombre = " ".join(partes[3:])  # El nombre es el resto de la línea
                alumnos.append((curp, nombre))

    return nombre_escuela, grado, grupo, alumnos


    
# Función para crear y guardar el DataFrame como archivo Excel
def crear_excel(nombre_escuela, grado, grupo, alumnos, ruta_salida):
    # Extraer CURPs y Nombres directamente de las tuplas
    curps = [alumno[0] for alumno in alumnos]  # CURP está en la posición 0 de la tupla
    nombres = [alumno[1] for alumno in alumnos]  # Nombre está en la posición 1 de la tupla
    
    df = pd.DataFrame({'CURP': curps, 'Nombre': nombres})
    nombre_archivo_excel = f"{nombre_escuela} {grado} {grupo}.xlsx"
    #print(f"Nombre: {nombre_archivo_excel}")
    df.to_excel(ruta_salida + nombre_archivo_excel, index=False)
    return ruta_salida + nombre_archivo_excel

"""
# Ejemplo con archivo individual

# Ruta al archivo PDF
ruta_pdf = "./PDFs/4B.pdf"
texto_pdf = extraer_texto_pdf(ruta_pdf)

# Procesar el texto extraído
nombre_escuela, grado, grupo, alumnos = procesar_texto(texto_pdf)

# Crear DataFrame y guardar en Excel
df = pd.DataFrame(alumnos, columns=['CURP', 'Nombre'])
nombre_archivo = f"{nombre_escuela} {grado} {grupo}.xlsx"
df.to_excel(nombre_archivo, index=False)

# Ruta de salida para el archivo Excel
ruta_salida = "./Excels/"  # Ajusta esto a la ruta deseada

# Crear y guardar el archivo Excel
archivo_excel = crear_excel(nombre_escuela, grado, grupo, alumnos, ruta_salida)
print(f"Archivo Excel creado: {archivo_excel}")
"""

# Define la ruta de la carpeta que contiene los PDFs
ruta_carpeta_pdfs = './PDFs/'
ruta_salida_excel = './Excels/'  # Ajusta esto a la ruta deseada

# Lista todos los archivos PDF en la carpeta especificada
archivos_pdf = [archivo for archivo in os.listdir(ruta_carpeta_pdfs) if archivo.endswith('.pdf')]

# Procesa cada archivo PDF
for archivo_pdf in archivos_pdf:
    ruta_completa_pdf = os.path.join(ruta_carpeta_pdfs, archivo_pdf)
    texto_pdf = extraer_texto_pdf(ruta_completa_pdf)
    
    # Procesa el texto extraído para obtener la información deseada
    nombre_escuela, grado, grupo, alumnos = procesar_texto(texto_pdf)
    
    # Crea y guarda el archivo Excel
    crear_excel(nombre_escuela, grado, grupo, alumnos, ruta_salida_excel)
    print(f"Archivo Excel creado para: {archivo_pdf}")