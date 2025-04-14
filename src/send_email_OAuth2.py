# Para enviar correo, permite probar si la rutina Enviar_Correo(), funciona correctamente.
# Para limpiar la pantalla
import os
#
import base64
#
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
#
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# Ruta al archivo JSON de credenciales
#CREDENTIALS_FILE = 'credentials.json'
CREDENTIALS_FILE = 'clave_Reclamos_Form_Web.json'
SCOPES = ['https://www.googleapis.com/auth/gmail.send']

# ===================================================================
# Class: Definimos los códigos de colores ANSI
# ===================================================================
class TextColor:
    RED = '\033[31m'
    GREEN = '\033[32m'
    YELLOW = '\033[33m'
    BLUE = '\033[34m'
    MAGENTA = '\033[35m'
    CYAN = '\033[36m'
    RESET = '\033[0m'  # Resetea el color al predeterminado
    # Ejemplo de uso
    #print(f"{TextColor.RED}Este texto es rojo.{TextColor.RESET}")


def authenticate_gmail():
    # Autenticar y obtener las credenciales OAuth2
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        # Guardar las credenciales en un archivo
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return creds

def enviar_correo(destinatario, asunto, cuerpo, archivo_adjunto):
    creds = authenticate_gmail()
    service = build('gmail', 'v1', credentials=creds)

    # Crear el mensaje de correo
    mensaje = MIMEMultipart()
    mensaje['to'] = destinatario
    mensaje['subject'] = asunto
    mensaje.attach(MIMEText(cuerpo, 'plain'))

    # Adjuntar el archivo
    if archivo_adjunto:
        with open(archivo_adjunto, 'rb') as adjunto:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(adjunto.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename= {os.path.basename(archivo_adjunto)}')
            mensaje.attach(part)

    # Codificar el mensaje en base64
    mensaje_bytes = mensaje.as_string().encode('utf-8')
    mensaje_base64 = base64.urlsafe_b64encode(mensaje_bytes).decode('utf-8')

    # Enviar el correo
    send_message = {'raw': mensaje_base64}
    try:
        service.users().messages().send(userId='me', body=send_message).execute()
        print("Correo enviado con éxito.")
    except Exception as e:
        print(f"Error al enviar correo: {e}")


#----------Datos Para enviar Correo --------------------------------------------------
nombre_archivo = 'RECLAMOS LUZ_241104_1050.docx'
#remitente ='enrecat@catamarca.gov.ar'
#password ='enrecat16'

'''
Ir a la cuenta de Google, Seguridad y Acceso de aplicaciones menos seguras. 
Hay que habilitar esta opcion
'''


#==============================================================================================================================
# Rutina para enviar correo
# 
"""
def Enviar_Correo(destinatario, asunto, cuerpo, archivo_adjunto, remitente, password):
    # Envia Correo con elemento adjunto
    # Obtener directorio actual del script
    directorio_script = os.path.dirname(os.path.abspath(__file__))
    
    # Subir un nivel hacia la carpeta que contiene 'Outputs'
    directorio_padre = os.path.dirname(directorio_script)   
    
    # Construir ruta completa al archivo
    ruta_archivo = os.path.join(directorio_padre, 'Outputs', archivo_adjunto)
    
    # Configuración del mensaje
    mensaje = MIMEMultipart()
    mensaje['From'] = remitente
    mensaje['To'] = destinatario
    mensaje['Subject'] = asunto

    # Adjuntar cuerpo del mensaje
    mensaje.attach(MIMEText(cuerpo, 'plain'))

    # Ruta completa del archivo adjunto
    #ruta_archivo = os.path.join(OUTPUT_PATH, archivo_adjunto)

    # Adjuntar archivo
    with open(ruta_archivo, 'rb') as adjunto:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(adjunto.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename= {archivo_adjunto}')
        mensaje.attach(part)

    # Establecer conexión con el servidor SMTP
    tu_servidor_smtp = 'smtp.gmail.com'
    smtp_port = 587 # Puerto seguro TLS
    servidor_smtp = smtplib.SMTP(host=tu_servidor_smtp, port=smtp_port)
    
    servidor_smtp.starttls() # Habilitar seguridad TLS

    text = mensaje.as_string()
    
    # Autenticación
    servidor_smtp.login(remitente, password)
    
    # Envío del correo
    servidor_smtp.sendmail(remitente, destinatario, mensaje.as_string())

    # Cerrar conexión
    servidor_smtp.quit()
"""

def main():
    try:
        global nombre_archivo
        #----------Datos Para enviar Correo --------------------------------------------------
        destinatario ='daguirreie@yahoo.com.ar'
        asunto ='RECLAMOS'
        cuerpo ='Hola, adjunto te envio RECLAMO al dia de la Fecha.'
        archivo_adjunto = nombre_archivo
        remitente ='enrecat@catamarca.gov.ar'
        password ='enrecat16'
        enviar_correo(destinatario, asunto, cuerpo, archivo_adjunto)
        #--------------------------------------------------------------------------------------
    except Exception as e:
        print(f'{TextColor.RED}Error:{TextColor.RESET} {e}')    

if __name__ == '__main__':
    # FUNCIONA BIEN ESTE SCRIPT.
    main()
    print(f"{TextColor.MAGENTA}<<----- FINAL PROGRAM ---->>{TextColor.RESET}")
    exit(0)
