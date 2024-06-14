# ---------- SE CONECTA USANDO oauth2 ---------------------------------------------------------------
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from google.oauth2 import service_account
#
from pathlib import Path
from googleapiclient.errors import HttpError
#
import os
# Para word
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
# Para
import shutil
# Para leer dato por consola
import sys
# Para Depurar un programa
import pdb
# Para sacar fecha de hoy
from datetime import datetime
# Para usar otra manera de crear documentos de word
from docx import Document

# CONFIGURACION DE USUARIOS
# ------- PARA EL GOOGLE SHEET
# Alcances necesarios para acceder a la API de Google Sheets
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# Escribe aquí el ID de tu documento:
# ID o URL de la hoja de cálculo y el Rango
SPREADSHEET_ID = '1FLWBfOe_ZKTMOviNBR35aC4CkHDGOwN2djyBd-rO0Js'

global SHEET_NAME
SHEET_NAME='' # Aqui voy a tener con que hoja trabajo. si es de LUZ o es de AGUA
SHEET_NAME_PRUEBA = 'ReclamosRes055-20'
SHEET_NAME_REC_LUZ = 'RtaFormRecLuz'
SHEET_NAME_REC_AGUA = 'RtaFormRecAgua'

# -------- PARA EL WORD
# Ruta al fichero Excel
#EXCEL_PATH= '.\Inputs\People_Data.xlsx'

# Ruta de Salida
OUTPUT_PATH= '.\Outputs'

# Ruta plantillas o Templates. Ficheros word
global WORD_TEMPLATE
WORD_TEMPLATE=''
ES_WORD_TPL_PATH='.\Inputs\Templates\WordTemplate_ES.docx'
EN_WORD_TPL_PATH='.\Inputs\Templates\WordTemplate_EN.docx'
WORD_TPL_PRUEBA1='.\Inputs\Templates\WordTemplate_Prueba1.docx'
WORD_TPL_PRUEBA_L2='.\Inputs\Templates\TemplateRECLAMOS_LUZ2.docx'
WORD_TPL_PRUEBA_A2='.\Inputs\Templates\TemplateRECLAMOS_AGUA2.docx'

# Ruta de Imagenes
IMAGE_PATH='.\Inputs\Images'

#==============================================================================================================================
# Limpia pantalla
def clear_screen():
    # Detecta el sistema operativo
    if os.name == 'nt':  # Para Windows
        os.system('cls')
    else:  # Para Unix/Linux/MacOS
        os.system('clear')
#------------------------------------------------------------------------------------------------------------------------------

#==============================================================================================================================
# Defino si es un Reclamo de LUZ o de AGUA.
def TipoReclamo(reclamo):
    global SHEET_NAME
    global WORD_TEMPLATE
    print(f"Tipo de Reclamo: {reclamo}")
    if reclamo == 'LUZ':
        # Es un reclamo de LUZ
        SHEET_NAME = SHEET_NAME_REC_LUZ
        WORD_TEMPLATE = WORD_TPL_PRUEBA_L2
        # Falta definir la planilla
    else:
        # Es un reclamo de AGUA
        SHEET_NAME = SHEET_NAME_REC_AGUA
        WORD_TEMPLATE = WORD_TPL_PRUEBA_A2
        #SHEET_NAME=SHEET_NAME_PRUEBA
        #exit(0)
    print(f"\033[34m HOJA SELECCIONADA: {SHEET_NAME}\033[0m")
    print(f"\033[34m PLANTILLA USADA: {WORD_TEMPLATE}\033[0m")
#------------------------------------------------------------------------------------------------------------------------------

#==============================================================================================================================
# Encuentra la PRIMERA FILA no leída (donde la primera columna no es verde)
def find_first_unread_row(rows):
    for i, row in enumerate(rows):
        cell = row['values'][0]
        if 'effectiveFormat' in cell:
            background = cell['effectiveFormat']['backgroundColor']
            # Verifica si el color de la celda es VERDE (0,1,0)
            if not (background.get('red', 0) == 0 and background.get('green', 0) == 1 and background.get('blue', 0) == 0):
                # No es VERDE la celda
                print("\033[34m Primera Fila Sin Leer:\033[0m",i+2)
                return i + 2
        else:
            #
            #return i + 2
            return None
                 
    #print("\033[34m Primera Fila:\033[0m",i+2)
    return None
#------------------------------------------------------------------------------------------------------------------------------

#==============================================================================================================================
# Devuelve la CANTIDAD de FILAS NO leídas (donde la primera columna no es verde)
def find_cant_unread_row(rows):
    cont_Filas_No_Verdes = 0
    for i, row in enumerate(rows):
        cell = row['values'][0]
        if 'effectiveFormat' in cell:
            background = cell['effectiveFormat']['backgroundColor']
            # Verifica si el color es verde
            if not (background.get('red', 0) == 0 and background.get('green', 0) == 1 and background.get('blue', 0) == 0):
                # No es VERDE la celda
                cont_Filas_No_Verdes = cont_Filas_No_Verdes + 1
                if (background.get('red', 0) == 1 and background.get('green', 0) == 0 and background.get('blue', 0) == 0):
                    # Si la celda es de COLOR ROJO sale.
                    return cont_Filas_No_Verdes
        else:
            #
            return cont_Filas_No_Verdes
    
    print("\033[34m Cant. Filas No Verdes:\033[0m",cont_Filas_No_Verdes)
    return cont_Filas_No_Verdes
#------------------------------------------------------------------------------------------------------------------------------

#==============================================================================================================================
# 
def print_row_data(row):
    """
    Imprime los datos de una fila.
    :param row: Lista con los datos de la fila.
    """
    print(', '.join(row))
    print('-' * 80)  # Línea separadora entre filas
#------------------------------------------------------------------------------------------------------------------------------

#==============================================================================================================================
# Rutina para eliminar y crear carpeta
def EliminarCrearCarpetas(path):
    #Verificar si la carpeta existe y elimninarla
    if(os.path.exists(path)):
        shutil.rmtree(path)
        
    # Crear carpeta
    os.mkdir(OUTPUT_PATH)
#------------------------------------------------------------------------------------------------------------------------------

#==============================================================================================================================
# Rutina para crear un fichero word para cada persona 
# ESTA FORMA SI USA PLANTILLA PRE DEFINIDA. 
def CrearWordPersonas(df_pers):
    # Iteramos sobre cada Persona
    for r_idx, r_val in enumerate(df_pers):
        # Cargar plantilla
        l_tpl=WORD_TEMPLATE   # Plantilla o Template que se va a usar.
        '''
        if (r_val['Idioma'] == 'ES'):
            l_tpl=ES_WORD_TPL_PATH
        elif (r_val['Idioma'] == 'EN'):
            l_tpl=EN_WORD_TPL_PATH
        '''
        # Procesamos la plantilla
        docx_tpl=DocxTemplate(l_tpl)

        # Añadir imagen grafico circular y de barra
        #img_path = IMAGE_PATH + '\\' + r_val['Imagen']
        #img = InlineImage(docx_tpl, img_path, height=Mm(15))

        # Crear contexto
        # word : Google Sheet
        
        context = {
            'name': r_val['Nombre'],
            'surname1': r_val['Apellido'],
            'surname2': r_val['Telefono de Contacto'], # En este me da error
            'edad': r_val['Correo Electrónico'],
            #'picture': img,
        }
        # Crear el contexto (ChatGPT)
        #contexto = {'items': [{'indice': i, 'valor': v} for i, v in enumerate(df_pers)]}

        # Verificar el contexto
        print("Contexto:", context)

        # Renderizamos usando el contexto creado
        docx_tpl.render(context)

        # Guardamos el documento
        nombre_doc = 'Documento_' + r_val['Apellido'].upper() + '_' + r_val['Nombre'] + '.docx'
        '''
        if(pd.notna(r_val['Apellido2'])):
            nombre_doc = 'Documento_' + r_val['Apellido1'].upper() + '_' + r_val['Apellido2'].upper() + '_' + r_val['Nombre'] + '.docx'
        else:
            nombre_doc = 'Documento_' + r_val['Apellido1'].upper() + '_' + r_val['Nombre'] + '.docx'
        docx_tpl.save(OUTPUT_PATH + '\\' + nombre_doc)
        '''
        docx_tpl.save(OUTPUT_PATH + '\\' + nombre_doc)
#------------------------------------------------------------------------------------------------------------------------------

#==============================================================================================================================
# Rutina para crear un fichero word para TODAS las personas 
# ESTA FORMA SI USA PLANTILLA PRE DEFINIDA. 
def crea_documento_unico(datos_para_diccionario):
    try:
        # Convertir la lista a una lista de diccionarios
        #lista_diccionarios = [{"nombre": item[0], "edad": item[1]} for item in datos]
        
        # Cargar la plantilla
        doc = DocxTemplate(WORD_TEMPLATE)
        
        # Crear el contexto con todos los datos
        contexto = {'items': datos_para_diccionario}
        
        # Verificar el contexto
        print("Contexto:", contexto)
        # Renderizar el documento con el contexto
        doc.render(contexto)
        
        # Obtener la fecha actual
        fecha_actual = datetime.now()
        # Extraer día, mes y año
        dia = fecha_actual.day
        mes = fecha_actual.month
        anio = fecha_actual.year
        
        # Formatear el nombre del archivo
        nombre_archivo = f"RECLAMO_{anio}{mes:02d}{dia:02d}.docx"

        if SHEET_NAME == SHEET_NAME_REC_LUZ:
            # Formatear el nombre del archivo
            nombre_doc = f"RECLAMOS_LUZ_{anio}{mes:02d}{dia:02d}.docx"

        elif SHEET_NAME == SHEET_NAME_REC_AGUA:
            # Formatear el nombre del archivo
            nombre_doc = f"RECLAMOS_AGUA_{anio}{mes:02d}{dia:02d}.docx"

        # Guardar el documento
        doc.save(OUTPUT_PATH + '\\' + nombre_doc)
        
        print(f"Documento guardado como: {nombre_doc}")
    
    except Exception as e:
        print("Ocurrió un error:", e)
#------------------------------------------------------------------------------------------------------------------------------

#==============================================================================================================================
# Rutina para crear un fichero word para TODAS las personas (OTRA MANERA)
# ESTA FORMA NO USA PLANTILLA PRE DEFINIDA. 
def OtraFormaCrearWord(datos_para_diccionario):
    # Obtener la fecha actual
    fecha_actual = datetime.now()
    anio = fecha_actual.year
    mes = fecha_actual.month
    dia = fecha_actual.day

    # Formatear el nombre del archivo
    nombre_archivo = f"RECLAMO_{anio}{mes:02d}{dia:02d}.docx"

    # Crear un documento de Word
    documento = Document()

    # Definir contador de hojas
    contador_hojas = 1

    # Iterar sobre cada diccionario en la lista y agregarlo al documento
    for indice, diccionario in enumerate(datos_para_diccionario):
        # Agregar el contenido del diccionario al documento
        documento.add_paragraph(f"RECLAMO DE AGUA NRO: {contador_hojas}")
        documento.add_paragraph(f"MARCA: {diccionario['Marca_Temporal']}")
        documento.add_paragraph(f"Apellido: {diccionario['Apellido']}")
        documento.add_paragraph(f"NOMBRE: {diccionario['Nombre']}")
        documento.add_paragraph(f"DNI: {diccionario['DNI']}")
        documento.add_paragraph(f"NRO. TEL: {diccionario['Nro_de_Telefono']}")
        documento.add_paragraph(f"CORREO: {diccionario['E_Mail']}")
        documento.add_paragraph(f"DOMICILIO: {diccionario['Domicilio']}")
        documento.add_paragraph(f"NRO. SUMINISTRO: {diccionario['Nro_de_Suministro']}")
        documento.add_paragraph(f"DESCRIPCION: {diccionario['Descripcion_Reclamo']}")

        # Agregar un salto de página después de cada elemento
        documento.add_page_break()

        # Incrementar el contador de hojas
        contador_hojas += 1

    # Guardar el documento
    documento.save(nombre_archivo)

    print(f"Archivo '{nombre_archivo}' creado exitosamente.")    
#------------------------------------------------------------------------------------------------------------------------------


#==============================================================================================================================
# RUTINA PRINCIPAL
def main():
    global SHEET_NAME

    # Ruta al archivo de credenciales .JSON
    current_dir = Path(__file__).parent
    KEY=current_dir/'clave_Reclamos_Form_Web.json'
    try:
        # Autenticación y acceso a la hoja de cálculo
        creds = None
        creds = service_account.Credentials.from_service_account_file(KEY, scopes=SCOPES)
        service = build('sheets', 'v4', credentials=creds)

        # Llamada a la api
        sheet = service.spreadsheets()
        #pdb.set_trace()
        # Obtén el rango de datos existente en la hoja, por medio de la columna A
        result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=f'{SHEET_NAME}!A:A').execute()
        '''
        result=
        range: NombreHoja!A1:A139
        majorDimension:ROWS
        values: [[Titulo Celda de A1],[],[],[],.......[valor de la celda A139 en este caso]] (es una lista de lista)
        '''
        print("\033[34m Resul:\033[0m") # Imprime en color azul
        print(result) # Imprime el diccionario
        
        rows = result.get('values', []) # Aqui solo son las filas (rows) de la primera columna (A). Hasta la ultima fila que tiene valor
        print(f"\033[34m {len(rows)} : Filas (rows) Recuperadas.\033[0m")
        # Saco la cantidad de filas que tienen datos, per no sabemos si estan pintadas
        num_rows = len(result.get('values', [])) - 1 # Descuento la cabecera
            
        '''
        #--chatGPT-------------------------------------------------------------------------------
        # Obtiene los detalles de la hoja de cálculo
        spreadsheet = sheet.get(spreadsheetId=SPREADSHEET_ID).execute()

        # Busca el sheetId correspondiente al SHEET_NAME
        sheet_id = None
        for sheet in spreadsheet['sheets']:
            if sheet['properties']['title'] == SHEET_NAME:
                sheet_id = sheet['properties']['sheetId']
                break
        #---------------------------------------------------------------------------------
        '''

        '''
        #---------------------------------------------------------------------------------
        # Imprime los títulos de las hojas y sus IDs
        for sheet in spreadsheet['sheets']:
            title = sheet['properties']['title']
            sheet_id = sheet['properties']['sheetId']
            print(f"Sheet title: {title}, Sheet ID: {sheet_id}")
        #---------------------------------------------------------------------------------
        '''

        if num_rows == 0:
            print(f"\033[34m No hay datos en la hoja: {SHEET_NAME}\033[0m")
        else:
            # Rango sin encabezado
            RANGE = f'{SHEET_NAME}!A2:A{num_rows+1}'  # Rango basado en el número de filas con datos, sin tener en cuenta si estan pintadas o no.
           
            #===========================================
            # Obtén los datos y el formato de la hoja
            #-------------------------------------------------------------------------------------------------------------------------------------------------
            # 1 Forma: El RANGE es sin el encabezado.
            result = sheet.get(spreadsheetId=SPREADSHEET_ID, ranges=RANGE, fields='sheets(data.rowData.values.effectiveFormat)').execute() # Con chatGPT. OK
            rows = result['sheets'][0]['data'][0]['rowData'] # Con chatGPT. OK
            # rowDat:{values:[{..}]}
            #-------------------------------------------------------------------------------------------------------------------------------------------------
            
            '''
            #-------------------------------------------------------------------------------------------------------------------------------------------------
            # De esta manera no especifico con que hoja (sheet) quiero trabajar. Por defecto toma la hoja 1.
            # 2 Forma:
            result = sheet.get(spreadsheetId=SPREADSHEET_ID, fields='sheets(data.rowData.values.effectiveFormat)').execute() # Con chatBlackbox. Me toma la sheet1. OK
            print("\033[34m Result: \033[0m")
            print(result)
            rows = result.get('sheets')[0].get('data')[0].get('rowData') # Con chatblackbox. OK
            print("\033[34m rows: \033[0m")
            print(rows)
            #-------------------------------------------------------------------------------------------------------------------------------------------------
            '''
            '''
            #-------------------------------------------------------------------------------------------------------------------------------------------------
            # 
            3 Forma:
            result = service.spreadsheets().values().batchGet(spreadsheetId=SPREADSHEET_ID, ranges=RANGE).execute() # De chatBlackbox. OK
            rows = result['sheets'][0]['data'][0]['rowData'] # No OK
            #-------------------------------------------------------------------------------------------------------------------------------------------------
            '''
            
            # Busco la primera fila que no esta pintada de VERDE, del rango (RANGE) especificado.
            first_unread_row = find_first_unread_row(rows)
            if not (first_unread_row == None):
                cant_unread_row = find_cant_unread_row(rows) # Cantidad de filas no leidas (NO VERDE)
                last_unread_row = (cant_unread_row + first_unread_row) - 1 # Saco la ultima fila del Rango que NO tiene VERDE.
            else:
                first_unread_row = 0 # Le pongo valor cero para decir que no hay filas nuevas.
            
            if first_unread_row:
                # Procesa el registro en la primera fila no leída
                print("\033[34m Procesando desde la fila: \033[0m", first_unread_row)
                # Este rango es donde esta mi informacion, nueva.
                if(SHEET_NAME == SHEET_NAME_REC_LUZ ):               
                    range_to_read = f'{SHEET_NAME}!A{first_unread_row}:J{last_unread_row}'  # Ajusta el rango según sea necesario
                elif(SHEET_NAME == SHEET_NAME_REC_AGUA):
                    range_to_read = f'{SHEET_NAME}!A{first_unread_row}:I{last_unread_row}'  # Ajusta el rango según sea necesario
                
                # ----CHATGPT
                rango_base=range_to_read.split('!')[1]  # Esto te dará 'CELDAx:CELDAy' CELDA=Letra, x e y numeros
                # Extraer el número de fila base
                fila_base = int(rango_base.split(':')[0][1:])  # Esto te dará nro de fila 
                # ---------------------------------------------------------------

                record = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=range_to_read).execute()
                print("\033[34mRECORD: \033[0m") # Color AZUL
                # Obtengo la LISTA.
                datos_lista = record.get('values',[])
                
                # La primera fila contiene los nombres de las columnas
                columns = datos_lista[0][1] # apunta a la segunda columna en este caso 'Apellido' [Fila][Columna]
                valor = datos_lista[1:] # Me devuelve todos los datos a partir de la fila 1.
                
                '''
                # --- Chat GPT -----------------------------------------------
                # Remueve el carácter de nueva línea del quinto campo
                campo_corregido = datos_lista[8].replace("\n", "")

                # Crea una nueva lista con el campo corregido
                lista_corregida = datos_lista[:8] + [campo_corregido]

                # Imprime la nueva lista corregida
                print(lista_corregida)
                #-------------------------------------------------------------
                '''

                # --- Chat GPT -----------------------------------------------
                # Especifica el índice del campo que deseas corregir (en este caso, el campo 5 tiene índice 4)
                indice_campo_a_corregir = 8

                # Inicializa una nueva lista para almacenar las listas corregidas
                lista_corregida = []
                # Imprime el tipo de 
                print(type(lista_corregida))

                # Recorre cada sublista en la lista de listas
                for sublista in datos_lista:
                    # Verifica si el campo a corregir es una cadena antes de intentar usar replace
                    if isinstance(sublista[indice_campo_a_corregir], str):
                        # Corrige el campo eliminando el carácter de nueva línea
                        sublista[indice_campo_a_corregir] = sublista[indice_campo_a_corregir].replace("\n", "")
                    
                    # Añade la sublista corregida a la nueva lista
                    lista_corregida.append(sublista)

                # Imprime la nueva lista de listas corregida
                print(lista_corregida)

                datos_lista = lista_corregida
                # Imprime el tipo de 
                print(type(datos_lista))

                #------------------------------------------------------------- 

                if not datos_lista:
                    print('No data found.')
                else:
                    # Itera sobre las filas y las imprime
                    for row in datos_lista:
                        print_row_data(row)
                    
                #pdb.set_trace()
                if(SHEET_NAME == SHEET_NAME_REC_LUZ ):               
                    # Convertir la lista en una lista de diccionarios
                    datos_lista_diccionario = [{'Marca_Temporal': item[0], 
                                        'Apellido': item[1], 
                                        'Nombre' : item[2], 
                                        'DNI' : item[3],
                                        'Nro_de_Telefono' : item[4],
                                        'E_Mail' : item[5],
                                        'Domicilio' : item[6],
                                        'Nro_de_Suministro' : item[7],
                                        'Tipo_de_Reclamo' : item[8],
                                        'Descripcion_Reclamo' : item[9]} for item in datos_lista]
                elif(SHEET_NAME == SHEET_NAME_REC_AGUA):
                    # Convertir la lista en una lista de diccionarios
                    datos_lista_diccionario = [{'Marca_Temporal': item[0], 
                                        'Apellido': item[1], 
                                        'Nombre' : item[2], 
                                        'DNI' : item[3], 
                                        'Nro_de_Telefono' : item[4],
                                        'E_Mail' : item[5],
                                        'Domicilio' : item[6],
                                        'Nro_de_Suministro' : item[7],
                                        'Descripcion_Reclamo' : item[8]} for item in datos_lista]

                # Imprime el tipo de datos_diccionario
                print(type(datos_lista_diccionario))
                
                #--- ChatGPT -----------------------------------------------------------------------
                # Recorre la lista de diccionario e imprime solo el campo especifico del diccionario
                # Define el campo específico que quieres imprimir
                campo_especifico = "Descripcion_Reclamo"
                # Verificar si datos_diccionario es realmente un diccionario
                
                if isinstance(datos_lista_diccionario, list):
                    for diccionario in datos_lista_diccionario:
                        # Verifica si el campo específico existe en el diccionario
                        if campo_especifico in diccionario:
                            print("")
                            print(diccionario[campo_especifico])
                        else:
                            print(f"{campo_especifico} no encontrado en el diccionario: {diccionario}")                    
                
                    '''
                    for clave, sub_diccionario in datos_diccionario.items():
                        if campo_especifico in sub_diccionario:
                            print(sub_diccionario[campo_especifico])
                        else:
                            print(f"{campo_especifico} no encontrado en {clave}")
                    '''
                else:
                    print("datos_lista_diccionario, no es una LISTA ")
                #---------------------------------------------------------------------------------------
                
                '''
                # Convierte los datos a un DataFrame de Pandas
                if not datos_lista:
                    print('No data found.')
                else:
                    # Asume que la primera fila de values contiene los nombres de las columnas
                    #df = pd.DataFrame(datos[1:], columns=datos[0])
                    # Muestra el DataFrame
                    #print(df)
                '''
                
                EliminarCrearCarpetas(OUTPUT_PATH)
                #CrearWordPersonas(datos_diccionario) # Crea una hoja de word por reclamo.
                #crea_documento_unico(datos_lista_diccionario) # Crea una hoja de word por multiples reclamos.
                OtraFormaCrearWord(datos_lista_diccionario) # Crea una hoja de word por multiples reclamos.
                # oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo
                '''
                # --- SACO sheet_id (chatGPT) ------------------------------------------------------------------------------
                # Obtiene los detalles de la hoja de cálculo. Saco el sheet_id
                spreadsheet = sheet.get(spreadsheetId=SPREADSHEET_ID).execute()
                # Busca el sheetId correspondiente al SHEET_NAME
                sheet_id = None
                for sheet in spreadsheet['sheets']:
                    if sheet['properties']['title'] == SHEET_NAME:
                        sheet_id = sheet['properties']['sheetId']
                        break
                #---------------------------------------------------------------------------------

                # Marca la fila como leída, pintando la primera celda de verde
                # Tener en cuenta que la primer fila es la nro 0
                star_Fila = (first_unread_row - 1)
                end_Fila = (star_Fila + cant_unread_row)
                requests = [{
                    'updateCells': {
                        'range': {
                            'sheetId': sheet_id,    # Id de la hoja con la cual estamos trabajando
                            'startRowIndex': star_Fila, # Inicio del rango de filas.
                            'endRowIndex': end_Fila,   # Para abarcar mas filas (es exclusive)
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
                        }] * cant_unread_row, # Multiplica la cantidad de filas que quiero pintar
                        'fields': 'userEnteredFormat.backgroundColor'
                    }
                }]
                body = {'requests': requests}

                # Ejecuta la solicitud batchUpdate
                response = service.spreadsheets().batchUpdate(spreadsheetId=SPREADSHEET_ID, body=body).execute()
                #response = sheet.batchUpdate(spreadsheetId=SPREADSHEET_ID, body=body).execute() # Tengo error aqui
                print(f"\033[32m {cant_unread_row} :Filas Marcadas como leídas. \033[0m")
                '''
                # oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo
            else:
                print("\033[34m No hay registros nuevos para leer. \033[0m")
            
    except HttpError as error:
        print(f"\033[31m Un Error ha ocurrido: \033[0m {error}")
        return error

# /////////////////////////////////////////////////////////////////////////
if __name__ == '__main__':
    clear_screen()
    '''
    # Verifica si se pasó algún argumento
    # sys.arg[0] = Tiene el Nombre del script
    pdb.set_trace()
    if len(sys.argv) > 1:
        # Se paso argumento
        dato = sys.argv[1]
        print(f"Dato recibido: {dato}")
        
        # Aquí puedes tomar decisiones basadas en el valor de 'dato'
        if dato == "LUZ":
            print(f"Has seleccionado RECLAMO DE: {dato}")
            TipoReclamo(dato)
            main()
        elif dato == "AGUA":
            print(f"Has seleccionado RECLAMO DE: {dato}")
            TipoReclamo(dato)
            main()
        else:
            print("Opción no reconocida")
    else:
        print("No se proporcionó ningún dato - Fin del Programa.")
    
    '''
    TipoReclamo('AGUA')
    main()
    print("\033[35m ----- FINAL PROGRAM ---- \033[0m")
    exit(0)
