import os
import pandas as pd
from docx import Document
import logging

# Configurar el sistema de logs para registrar errores
logging.basicConfig(filename='app.log', level=logging.DEBUG)

# Función para cargar los datos desde un archivo Excel
def cargar_datos_excel(archivo_excel):
    """
    Carga datos desde un archivo Excel utilizando pandas.

    Parámetros:
    archivo_excel (str): Ruta del archivo Excel que contiene los datos.

    Retorna:
    pd.DataFrame: DataFrame con los datos del archivo Excel si se carga correctamente.
    None: Si ocurre un error al cargar el archivo (por ejemplo, si no se encuentra).
    """
    try:
        # Intentar cargar los datos desde el archivo Excel
        return pd.read_excel(archivo_excel)
    except FileNotFoundError:
        # Manejo de error si el archivo no se encuentra
        print(f"El archivo '{archivo_excel}' no se encuentra. Verifica la ruta o el nombre.")
        logging.error(f"El archivo '{archivo_excel}' no se encuentra. Verifica la ruta o el nombre.")
        return None

# Función para generar informes personalizados a partir de una plantilla de Word
def generar_informe(nombre, matematicas, ciencias, historia, plantilla):
    """
    Genera un informe de notas personalizado para un alumno, basado en una plantilla de Word.

    Parámetros:
    nombre (str): Nombre del alumno.
    matematicas (float): Nota en Matemáticas.
    ciencias (float): Nota en Ciencias.
    historia (float): Nota en Historia.
    plantilla (str): Ruta de la plantilla de Word.

    Retorna:
    str: Nombre del archivo del informe generado si se guarda correctamente.
    None: Si ocurre un error durante la generación del informe.
    """
    # Calcular el promedio de las notas
    promedio = (matematicas + ciencias + historia) / 3

    # Diccionario con los datos que serán reemplazados en la plantilla
    datos = {
        '{ALUMNO}': nombre,
        '{MATEMÁTICAS}': str(matematicas),
        '{CIENCIAS}': str(ciencias),
        '{HISTORIA}': str(historia),
        '{PROMEDIO}': f'{promedio:.2f}'  # Promedio formateado con dos decimales
    }

    try:
        # Cargar la plantilla de Word
        doc = Document(plantilla)

        # Reemplazar los campos de la plantilla con los valores reales del alumno
        for parrafo in doc.paragraphs:
            for campo, valor in datos.items():
                if campo in parrafo.text:
                    parrafo.text = parrafo.text.replace(campo, valor)

        # Guardar el archivo de informe personalizado en la carpeta 'uploads'
        nombre_archivo = f'nota_{nombre}.docx'
        ruta_archivo = os.path.join('uploads', nombre_archivo)
        doc.save(ruta_archivo)  # Guardar el documento generado

        return nombre_archivo  # Devolver el nombre del archivo generado

    except Exception as e:
        # Manejo de errores durante la generación del informe
        print(f"Error al generar informe para {nombre}: {e}")
        logging.error(f"Error al generar informe para {nombre}: {e}")
        return None
