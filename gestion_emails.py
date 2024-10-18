import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import logging
from retrying import retry
from dotenv import load_dotenv

# Configurar el log de errores
logging.basicConfig(filename='app.log', level=logging.DEBUG)

# Cargar las variables de entorno desde el archivo .env
load_dotenv()

EMAIL_ADDRESS = os.getenv('EMAIL_ADDRESS')
EMAIL_PASSWORD = os.getenv('EMAIL_PASSWORD')

if not EMAIL_ADDRESS or not EMAIL_PASSWORD:
    print("La dirección de correo o la contraseña no se han definido en las variables de entorno.")
    logging.error("La dirección de correo o la contraseña no se han definido en las variables de entorno.")
    exit()

# Clase para gestionar el envío de correos electrónicos
class GestorCorreo:
    def __init__(self, remitente, password):
        self.remitente = remitente
        self.password = password

    @retry(stop_max_attempt_number=3, wait_fixed=2000)
    def enviar_correo(self, destinatario, asunto, cuerpo, archivo_adjunto):
        # Crear el mensaje
        mensaje = MIMEMultipart()
        mensaje['From'] = self.remitente
        mensaje['To'] = destinatario
        mensaje['Subject'] = asunto

        # Agregar el cuerpo del mensaje
        mensaje.attach(MIMEText(cuerpo, 'plain'))

        # Verificar si el archivo existe antes de intentar abrirlo
        if not os.path.isfile(archivo_adjunto):
            print(f"Archivo no encontrado: {archivo_adjunto}")
            logging.error(f"Archivo no encontrado: {archivo_adjunto}")
            return

        # Adjuntar el archivo
        try:
            with open(archivo_adjunto, "rb") as adjunto:
                parte = MIMEBase('application', 'vnd.openxmlformats-officedocument.wordprocessingml.document')
                parte.set_payload(adjunto.read())
                encoders.encode_base64(parte)
                filename = os.path.basename(archivo_adjunto)
                parte.add_header('Content-Disposition', 'attachment', filename=filename)
                mensaje.attach(parte)

            # Configurar el servidor de correo
            print(f"Intentando conectar al servidor SMTP para enviar correo a {destinatario}...")
            logging.debug(f"Intentando conectar al servidor SMTP para enviar correo a {destinatario}...")

            try:
                with smtplib.SMTP_SSL('smtp.gmail.com', 465) as servidor:
                    servidor.set_debuglevel(1)  # Nivel de depuración del servidor SMTP
                    servidor.login(self.remitente, self.password)
                    servidor.sendmail(self.remitente, destinatario, mensaje.as_string())
                    print(f"Correo enviado a {destinatario}")
                    logging.info(f"Correo enviado a {destinatario}")
            except smtplib.SMTPAuthenticationError as auth_error:
                print(f"Error de autenticación al intentar enviar correo a {destinatario}: {auth_error}")
                logging.error(f"Error de autenticación al intentar enviar correo a {destinatario}: {auth_error}")
            except smtplib.SMTPException as smtp_error:
                print(f"Error SMTP al enviar correo a {destinatario}: {smtp_error}")
                logging.error(f"Error SMTP al enviar correo a {destinatario}: {smtp_error}")

        except Exception as e:
            logging.error(f"Error al enviar el correo a {destinatario}: {e}")
            print(f"Ha ocurrido un error al enviar el correo a {destinatario}. Revisa el archivo de logs.")
