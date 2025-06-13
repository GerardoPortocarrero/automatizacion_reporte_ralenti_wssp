import pandas as pd
import os

# Descargar archivos del correo
def download_mail_files(mails, MAIL_ITEM_CODE, project_address, conductores_ralenti):
    files_found = {} # {
                     #  9/06/2025: {'transportista': {}, 'ruta': {}},
                     #  10/06/2025: {'transportista': {}, 'ruta': {}},                     
                     # }

    for mail in mails:
        if mail.Class != MAIL_ITEM_CODE or mail.Attachments.Count == 0:
            continue

        subject_lower = mail.Subject.lower()
        string_date_mail = mail.ReceivedTime.strftime("%Y-%m-%d")

        for attachment in mail.Attachments:
            # Si no es un excel saltar
            if not attachment.FileName.endswith((".xlsx", ".xls")):
                continue
            
            # Si el correo tiene asunto carga diaria o venta perdida
            type_file = None
            sheet_name = None
            print(conductores_ralenti['mail_subject'].lower())
            print(subject_lower)
            print(conductores_ralenti['mail_subject'].lower() in subject_lower)
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
                print(df)
                # df_clean = df.dropna(how='all') # Elminar todos los NaN
                # if df.empty or df_clean.shape[0] == 0: # Si no hay datos eliminar y saltar
                #     os.remove(file_address)
                #     continue
            except Exception:
                # os.remove(file_address)
                continue

            # Guardar archivo para esa fecha
            if string_date_mail not in files_found:
                files_found[string_date_mail] = {}

            # Guardar tipo de archivo en esa fecha
            files_found[string_date_mail][type_file] = {
                'file_address': file_address,
                'received_time': string_date_mail
            }

            # Checar si ya tenemos ambos para esta fecha
            if conductores_ralenti['name'] in files_found[string_date_mail]:
                datos = files_found[string_date_mail]

                print(f'DATOS: {datos}')
                
                return (
                    datos[conductores_ralenti['name']]['file_address'], datos[conductores_ralenti['name']]['received_time'],                    
                )