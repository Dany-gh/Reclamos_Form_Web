# Para enviar correo
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# Para limpiar la pantalla
import os

'''
Ir a la cuenta de Google, Seguridad y Acceso de aplicaciones menos seguras. 
Hay que habilitar esta opcion
'''

nombre_archivo = 'RECLAMO_AGUA_20240617.docx'
#==============================================================================================================================
# Rutina para enviar correo
# 
def Enviar_Correo(destinatario, asunto, cuerpo, archivo_adjunto,remitente,password):
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

    text=mensaje.as_string()
    
    # Autenticación
    servidor_smtp.login(remitente, password)
    
    # Envío del correo
    servidor_smtp.sendmail(remitente, destinatario, mensaje.as_string())

    # Cerrar conexión
    servidor_smtp.quit()


def main():
    global nombre_archivo
    #----------Datos Para enviar Correo --------------------------------------------------
    destinatario ='daguirreie@yahoo.com.ar'
    asunto ='RECLAMOS'
    cuerpo ='Hola, adjunto te envio RECLAMO al dia de la Fecha.'
    archivo_adjunto =nombre_archivo
    remitente ='enrecat@catamarca.gov.ar'
    password ='enrecat16'
    Enviar_Correo(destinatario,asunto,cuerpo,archivo_adjunto,remitente,password)
    #--------------------------------------------------------------------------------------

if __name__ == '__main__':
    # FUNCIONA BIEN ESTE SCRIPT.
    main()
    print("\033[35m ----- FINAL PROGRAM ---- \033[0m")
    exit(0)
