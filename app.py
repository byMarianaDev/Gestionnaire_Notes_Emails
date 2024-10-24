from flask import Flask, request, render_template, send_file, jsonify
import os
from werkzeug.utils import secure_filename
from gestion_emails import GestorCorreo
from procesamiento_datos import cargar_datos_excel, generar_informe
import zipfile

app = Flask(__name__)

# Configuración de carpetas para subir archivos y generar informes
UPLOAD_FOLDER = 'uploads/'
GENERATED_REPORTS_FOLDER = 'generated_reports/'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['GENERATED_REPORTS_FOLDER'] = GENERATED_REPORTS_FOLDER

# Crear las carpetas si no existen para asegurar que estén disponibles
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)
if not os.path.exists(GENERATED_REPORTS_FOLDER):
    os.makedirs(GENERATED_REPORTS_FOLDER)

# Ruta principal que carga el formulario en la página principal
@app.route('/')
def index():
    return render_template('index.html')

# Función para crear un archivo ZIP que contenga todos los informes generados
def crear_backup_zip(carpeta_origen, archivo_zip_destino):
    with zipfile.ZipFile(archivo_zip_destino, 'w') as backup_zip:
        # Recorrer la carpeta para agregar solo los archivos de informes (que empiecen con 'nota_' y terminen con '.docx')
        for foldername, subfolders, filenames in os.walk(carpeta_origen):
            for filename in filenames:
                if filename.startswith('nota_') and filename.endswith('.docx'):
                    file_path = os.path.join(foldername, filename)
                    backup_zip.write(file_path, os.path.basename(file_path))  # Agregar solo el archivo, no su ruta completa

# Ruta para procesar los archivos subidos y generar los informes
@app.route('/procesar_archivos', methods=['POST'])
def procesar_archivos():
    # Verificar si los archivos están presentes en la solicitud
    if 'archivo_1' not in request.files or 'archivo_2' not in request.files:
        print("Archivos no presentes en la solicitud")
        return jsonify(success=False, message="Faltan archivos"), 400

    archivo_1 = request.files['archivo_1']  # Archivo Excel
    archivo_2 = request.files['archivo_2']  # Plantilla de Word

    # Verificar si los archivos fueron seleccionados
    if archivo_1.filename == '' or archivo_2.filename == '':
        print("Archivos no seleccionados")
        return jsonify(success=False, message="No se seleccionaron archivos"), 400

    # Guardar los archivos de manera segura
    ruta_archivo_1 = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(archivo_1.filename))
    ruta_archivo_2 = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(archivo_2.filename))

    print(f"Guardando archivo 1 en: {ruta_archivo_1}")
    print(f"Guardando archivo 2 en: {ruta_archivo_2}")

    try:
        archivo_1.save(ruta_archivo_1)
        archivo_2.save(ruta_archivo_2)
    except Exception as e:
        print(f"Error al guardar los archivos: {str(e)}")
        return jsonify(success=False, message=f"Error al guardar los archivos: {str(e)}"), 500

    print("Archivos guardados correctamente")

    # Cargar los datos desde el archivo Excel
    df = cargar_datos_excel(ruta_archivo_1)
    if df is None:
        return jsonify(success=False, message="Error al cargar el archivo Excel."), 400

    # Generar los informes personalizados basados en los datos del archivo Excel
    plantilla = ruta_archivo_2
    for index, fila in df.iterrows():
        nombre_informe = generar_informe(fila['Alumno'], fila['Matemáticas'], fila['Ciencias'], fila['Historia'], plantilla)
        if nombre_informe:
            # Mover los informes generados desde 'uploads' a 'generated_reports'
            ruta_informe = os.path.join(app.config['UPLOAD_FOLDER'], nombre_informe)
            nueva_ruta = os.path.join(app.config['GENERATED_REPORTS_FOLDER'], nombre_informe)
            os.rename(ruta_informe, nueva_ruta)  # Mover archivo generado

    # Crear archivo ZIP con todos los informes generados
    archivo_zip_destino = os.path.join(app.config['GENERATED_REPORTS_FOLDER'], 'backup_informes.zip')
    try:
        crear_backup_zip(app.config['GENERATED_REPORTS_FOLDER'], archivo_zip_destino)
    except Exception as e:
        return jsonify(success=False, message=f"Error al crear el archivo ZIP: {str(e)}"), 500

    return jsonify(success=True, message="Informes generados correctamente.")

# Ruta para descargar el archivo ZIP con los informes generados
@app.route('/descargar_backup')
def descargar_backup():
    archivo_zip = os.path.join(app.config['GENERATED_REPORTS_FOLDER'], 'backup_informes.zip')
    try:
        return send_file(archivo_zip, as_attachment=True)  # Descargar el archivo ZIP como adjunto
    except FileNotFoundError:
        return "El archivo ZIP no se encontró. Asegúrate de que se haya generado correctamente.", 404

# Ruta para enviar correos con los informes adjuntos a cada alumno
@app.route('/enviar_correos', methods=['POST'])
def enviar_correos():
    email = request.form['email']  # Correo del remitente
    password = request.form['password']  # Contraseña del remitente

    # Verificar que el archivo Excel con los datos esté disponible
    ruta_archivo_1 = os.path.join(app.config['UPLOAD_FOLDER'], 'notas_alumnos.xlsx')
    if not os.path.exists(ruta_archivo_1):
        return jsonify(success=False, message="El archivo Excel no fue encontrado."), 400

    # Cargar los datos del archivo Excel
    df = cargar_datos_excel(ruta_archivo_1)
    if df is None:
        return jsonify(success=False, message="Error al cargar el archivo Excel."), 400

    # Inicializar el gestor de correos
    gestor = GestorCorreo(email, password)

    # Enviar correo a cada alumno con su informe adjunto
    for index, fila in df.iterrows():
        archivo_informe = f"nota_{fila['Alumno']}.docx"
        ruta_informe = os.path.join(app.config['GENERATED_REPORTS_FOLDER'], archivo_informe)
        if not os.path.exists(ruta_informe):
            return jsonify(success=False, message=f"No se encontró el informe para {fila['Alumno']}"), 404
        asunto = f"Informe de Notas para {fila['Alumno']}"
        cuerpo = f"Estimado/a {fila['Alumno']}, adjunto encontrarás tu informe de notas."
        gestor.enviar_correo(fila['Correo'], asunto, cuerpo, ruta_informe)

    return jsonify(success=True, message="Correos enviados exitosamente.")

# Ruta para eliminar un archivo previamente subido
@app.route('/eliminar_archivo', methods=['POST'])
def eliminar_archivo():
    archivo = request.form.get('archivo')

    # Definir la ruta completa del archivo
    archivo_ruta = os.path.join(app.config['UPLOAD_FOLDER'], archivo)

    try:
        # Si el archivo existe, eliminarlo
        if os.path.exists(archivo_ruta):
            os.remove(archivo_ruta)
            return jsonify(success=True, message="Archivo eliminado correctamente")
        else:
            return jsonify(success=False, message="Archivo no encontrado"), 404
    except Exception as e:
        return jsonify(success=False, message=f"Error al eliminar el archivo: {str(e)}"), 500

# Iniciar la aplicación en modo debug
if __name__ == "__main__":
    app.run(debug=True)
