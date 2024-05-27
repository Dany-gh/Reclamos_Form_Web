# Prueba de conexion segun:
# https://www.youtube.com/watch?v=n0EkLvSOWc8
# NO puedo importar gspread sin antes instalarla
# Version es 6.1.2
# no puedo usar oauth2client.service_account, tengo que instalar oauth2client

import gspread
from oauth2client.service_account import ServiceAccountCredentials
from pathlib import Path

scope1_uno="https://spreadsheets.google.com/feeds"
scope1_dos='https://www.googleapis.com/auth/spreadsheets'
scope1_tres="https://www.googleapis.com/auth/drive"
scope1_cuatro="https://www.googleapis.com/auth/drive"

# Define los alcances necesarios
SCOPES = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets']
#SCOPES = [scope1_uno, scope1_dos]
#SCOPES=['https://www.googleapis.com/auth/spreadsheets']

current_dir = Path(__file__).parent
#print(current_dir)

# Ruta al archivo de credenciales JSON
#KEY=r'C:\Users\Daniel\Proyectos\Proyectos Python\Reclamos_Form_Web\src\clave_Reclamos_Form_Web.json'
KEY=current_dir/'clave_Reclamos_Form_Web.json'
#KEY=r'D:\Proyectos\Proyectos Python\Reclamos_Form_Web\src\clave_Reclamos_Form_Web.json'
#KEY = 'clave_Reclamos_Form_Web.json'

# Autenticación con las credenciales del archivo JSON
creds = ServiceAccountCredentials.from_json_keyfile_name(KEY, SCOPES)
client=gspread.authorize(creds)
print("Hola")

# A partir de aqui trabajamos con la hoja de calculo
try:
    # Abrimos segun ejemplo en youtube
    #sheet=client.open("BDENRE_RECLAMOS").sheet1
    #sheet.update_acell('B1','Bingo')
    
    # Abrir la hoja de cálculo por ID
    SPREADSHEET_ID = '1FLWBfOe_ZKTMOviNBR35aC4CkHDGOwN2djyBd-rO0Js'
    spreadsheet = client.open_by_key(SPREADSHEET_ID)
    # Seleccionar la hoja por nombre
    SHEET_NAME='RtaFormRecAgua'
    sheet = spreadsheet.worksheet(SHEET_NAME)
    print("**Conexión exitosa.")
    print(sheet.get_all_records())
    # Leer un rango específico de celdas
    values = sheet.get('A1:I1')
    # Imprimir los valores obtenidos
    for row in values:
        print(row)
except gspread.exceptions.SpreadsheetNotFound:
    print("**No se encontró la hoja de cálculo. Verifica el nombre y los permisos.")
except gspread.exceptions.GSpreadException as e:
    print(f"**Ocurrió un error con gspread: {e}")
except Exception as e:
    print(f"**Ocurrió un error general: {e}")

print("**FIN**.")
