from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from google.oauth2 import service_account

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
KEY = 'clave_Reclamos_Form_Web.json'
# Escribe aquí el ID de tu documento:
SPREADSHEET_ID = '1FLWBfOe_ZKTMOviNBR35aC4CkHDGOwN2djyBd-rO0Js'

creds = None
creds = service_account.Credentials.from_service_account_file(KEY, scopes=SCOPES)

service = build('sheets', 'v4', credentials=creds)
sheet = service.spreadsheets()

# Llamada a la api
result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range='RtaFormRecAgua!A1:I1').execute()
# Extraemos values del resultado
values = result.get('values',[])
print(values)