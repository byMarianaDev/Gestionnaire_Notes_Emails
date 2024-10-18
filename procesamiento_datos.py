import os
import pandas as pd
from docx import Document
import logging
from threading import Thread
from gestion_emails import GestorCorreo  # Importar la clase desde el archivo del colaborador 1

# Configurar el log de errores
logging.basicConfig(filename='app.log', level=logging.DEBUG)

# Función para cargar datos desde el archivo Excel
def cargar_datos_excel(archivo_excel):
    try:
        return pd.read_excel(archivo_excel)
    except FileNotFoundError:
        print(f"El archivo '{archivo_excel}' no se encuentra. Verifica la ruta o el nombre.")
        logging.error(f"El archivo '{archivo_excel}' no se encuentra. Verifica la ruta o el nombre.")
        exit()

# Función para generar el informe a partir de la plantilla
def generar_informe(nombre, matematicas, ciencias, historia, plantilla):
    promedio = (matematicas + ciencias + historia) / 3
    datos = {
        '{ALUMNO}': nombre,
        '{MATEMÁTICAS}': str(matematicas),
        '{CIENCIAS}': str(ciencias),
        '{HISTORIA}': str(historia),
        '{PROMEDIO}': f'{promedio:.2f}'
    }

    try:
        # Crear el documento a partir de la plantilla
        doc = Document(plantilla)
        for parrafo in doc.paragraphs:
            for campo, valor in datos.items():
                if campo in parrafo.text:
                    parrafo.text = parrafo.text.replace(campo, valor)

        # Guardar el archivo de informe personalizado
        nombre_archivo = f'nota_{nombre}.docx'
        doc.save(nombre_archivo)
        return nombre_archivo
    except FileNotFoundError:
        print(f"La plantilla '{plantilla}' no se encuentra. Verifica la ruta o el nombre.")
        logging.error(f"La plantilla '{plantilla}' no se encuentra. Verifica la ruta o el nombre.")
        exit()

# Cargar datos y generar informes
archivo_excel = 'notas_alumnos.xlsx'
plantilla = 'plantilla_notas.docx'

df = cargar_datos_excel(archivo_excel)

# Crear una instancia del gestor de correos (compartido con el colaborador 1)
gestor = GestorCorreo(os.getenv('EMAIL_ADDRESS'), os.getenv('EMAIL_PASSWORD'))

# Iterar sobre cada fila del archivo Excel para generar informes y enviar correos
def procesar_fila(fila):
    nombre = fila['Alumno']
    matematicas = fila['Matemáticas']
    ciencias = fila['Ciencias']
    historia = fila['Historia']
    destinatario = fila['Correo']

    archivo_informe = generar_informe(nombre, matematicas, ciencias, historia, plantilla)

    asunto = f'Informe de Notas para {nombre}'
    cuerpo = f"""Estimado/a {nombre}, adjunto encontrarás tu informe de notas.

Atentamente,
El Profesorado"""

    gestor.enviar_correo(destinatario, asunto, cuerpo, archivo_informe)

# Crear hilos para enviar correos de manera concurrente
hilos = []
for index, fila in df.iterrows():
    hilo = Thread(target=procesar_fila, args=(fila,))
    hilos.append(hilo)
    hilo.start()

# Esperar a que todos los hilos terminen
for hilo in hilos:
    hilo.join()

print("Correos enviados exitosamente.")
