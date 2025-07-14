import pandas as pd
import polars as pl
from datetime import timedelta
import os
import numpy as np

# Quedarme con la tabla principal
def get_main_table(df):
    # Tomar la fila 9 como nombres de columns
    df.columns = df.iloc[9]

    # Eliminar filas innecesarias
    df = df.iloc[10:].reset_index(drop=True)

    # Detectar el índice de la primera columna con nombre NaN
    if df.columns.hasnans:
        nan_index = df.columns.get_indexer_for([np.nan])[0]
        df = df.iloc[:, :nan_index]  # Cortar hasta antes del primer NaN

    # También elimina columnas vacías (por si acaso)
    df = df.dropna(axis=1, how='all')

    # Rellenar valores NaN hacia adelante
    df["LOCACION"] = df["LOCACION"].fillna(method='ffill')
    
    # Asegúrate de que no haya espacios y todo esté en mayúsculas
    df['LOCACION'] = df['LOCACION'].astype(str).str.strip().str.upper()
    
    # Filtrar por locacion
    df = df[df['LOCACION'] == 'PEDREGAL']

    return df

# Descargar archivos del correo
def download_mail_file(mails, MAIL_ITEM_CODE, project_address, conductores_ralenti):
    files_found = {} # {
                     #  9/06/2025: {'Conductores ralenti': {}},
                     #  10/06/2025: {'Conductores ralenti': {}},                     
                     # }

    for mail in mails:
        if mail.Class != MAIL_ITEM_CODE or mail.Attachments.Count == 0:
            continue

        subject_lower = mail.Subject.lower()
        mail_received_time = mail.ReceivedTime.strftime("%d-%m-%Y")
        mail_data_time = (mail.ReceivedTime - timedelta(days=1)).strftime("%d-%m-%Y") # Los datos son de una fecha anterior

        for attachment in mail.Attachments:
            # Si no es un excel saltar
            if not attachment.FileName.endswith((".xlsx", ".xls")):
                continue
            
            # Si el correo tiene asunto carga diaria o venta perdida
            type_file = None
            sheet_name = None
            
            if conductores_ralenti['mail_subject'].lower() in subject_lower:
                type_file = conductores_ralenti['name']
                sheet_name = conductores_ralenti['mail_sheet_name']
                
            else: # Si no coincide ningun asunto saltar
                continue
            
            # Guardar archivo excel con alguno de los asuntos
            file_name = attachment.FileName
            file_address = os.path.join(project_address, file_name)
            attachment.SaveAsFile(file_address)

            # Validar contenido
            try:
                df = pd.read_excel(file_address, sheet_name=sheet_name, header=None)                
                df = get_main_table(df)

                df['TIEMPO RALENTI SEGUNDOS.'] = pd.to_timedelta(df['TIEMPO RALENTI.'])
                tiempo_ralenti = int((df['TIEMPO RALENTI SEGUNDOS.'].sum()).total_seconds() // 60)
                
                if (df['RECORRIDO.'].sum() == 0) or (tiempo_ralenti == 0): # Si no hay datos eliminar y saltar
                    os.remove(file_address)
                    continue
            except Exception as e:
                print(e)
                os.remove(file_address)
                continue

            # Guardar archivo para esa fecha
            if mail_received_time not in files_found:
                files_found[mail_received_time] = {}

            # Guardar tipo de archivo en esa fecha
            files_found[mail_received_time][type_file] = {
                'file_address': file_address,
                'received_time': mail_received_time,
                'data_time': mail_data_time,
            }

            # Checar si ya tenemos ambos para esta fecha
            if conductores_ralenti['name'] in files_found[mail_received_time]:
                datos = files_found[mail_received_time]
                
                return (
                    df,
                    datos[conductores_ralenti['name']]['file_address'],
                    datos[conductores_ralenti['name']]['received_time'],
                    datos[conductores_ralenti['name']]['data_time'],
                )