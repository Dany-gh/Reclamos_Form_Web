# ---------- SE CONECTA USANDO oauth2 ---------------------------------------------------------------
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from google.oauth2 import service_account
from pathlib import Path
from googleapiclient.errors import HttpError


# Alcances necesarios para acceder a la API de Google Sheets
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# Escribe aquí el ID de tu documento:
# ID o URL de la hoja de cálculo y el Rango
SPREADSHEET_ID = '1FLWBfOe_ZKTMOviNBR35aC4CkHDGOwN2djyBd-rO0Js'
SHEET_NAME = 'ReclamosRes055-20'
#RANGE = 'RtaFormRecAgua!A1:A' # Rango dinámico para cubrir toda la columna A

def main():
    # Ruta al archivo de credenciales JSON
    current_dir = Path(__file__).parent
    KEY=current_dir/'clave_Reclamos_Form_Web.json'
    #KEY = 'clave_Reclamos_Form_Web.json'

    #try:
    # Autenticación y acceso a la hoja de cálculo
    creds = None
    creds = service_account.Credentials.from_service_account_file(KEY, scopes=SCOPES)
    service = build('sheets', 'v4', credentials=creds)

    # Llamada a la api
    sheet = service.spreadsheets()

    # Obtén el rango de datos existente en la hoja
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=f'{SHEET_NAME}!A:A').execute()
    print(result) # Imprime la lista

    rows = result.get('values', [])
    print(f"{len(rows)} rows retrieved")

    num_rows = len(result.get('values', []))

    if num_rows == 0:
        print("No hay datos en la hoja.")
    else:
        RANGE = f'{SHEET_NAME}!A1:A{num_rows}'  # Rango dinámico basado en el número de filas con datos

        #for row in rows:
        #    # Print
        #    print('%s, %s' % (row[0],row[1]))

        # Obtén los datos y el formato de la hoja
        result = sheet.get(spreadsheetId=SPREADSHEET_ID, range=RANGE, fields='sheets(data.rowData.values.effectiveFormat)').execute()
        rows = result['sheets'][0]['data'][0]['rowData']

        # Encuentra la primera fila no leída (donde la primera columna no es verde)
        def find_first_unread_row(rows):
            for i, row in enumerate(rows):
                cell = row['values'][0]
                if 'effectiveFormat' in cell:
                    background = cell['effectiveFormat']['backgroundColor']
                    # Verifica si el color es verde
                    if not (background.get('red', 0) == 0 and background.get('green', 0) == 1 and background.get('blue', 0) == 0):
                        return i + 1
                else:
                    return i + 1
            return None

        first_unread_row = find_first_unread_row(rows)

        if first_unread_row:
            # Procesa el registro en la primera fila no leída
            print("Procesando fila:", first_unread_row)
            range_to_read = f'{SHEET_NAME}!A{first_unread_row}:F{first_unread_row}'  # Ajusta el rango según sea necesario
            record = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=range_to_read).execute()
            print(record['values'])

            # Marca la fila como leída pintando la primera celda de verde
            requests = [{
                'updateCells': {
                    'range': {
                        'sheetId': 0,
                        'startRowIndex': first_unread_row - 1,
                        'endRowIndex': first_unread_row,
                        'startColumnIndex': 0,
                        'endColumnIndex': 1,
                    },
                    'rows': [{
                        'values': [{
                            'userEnteredFormat': {
                                'backgroundColor': {
                                    'red': 0,
                                    'green': 1,
                                    'blue': 0
                                }
                            }
                        }]
                    }],
                    'fields': 'userEnteredFormat.backgroundColor'
                }
            }]

            body = {
                'requests': requests
            }

            response = sheet.batchUpdate(spreadsheetId=SPREADSHEET_ID, body=body).execute()
            print('Fila marcada como leída.')
        else:
            print("No hay registros nuevos para leer.")


    #except HttpError as error:
    #    print(f"An error occurred: {error}")
    #    return error

if __name__ == '__main__':
    main()



# Extraemos values del resultado
#values = result.get('values',[])
#print(values)

# Imprime los valores obtenidos
#if not values:
#    print('No data found.')
#else:
#    for row in values:
#        print(row)

#print("FINAL")