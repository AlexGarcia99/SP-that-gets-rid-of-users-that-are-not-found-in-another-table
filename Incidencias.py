from datetime import datetime, timedelta
import re
import pyodbc
import pandas as pd
import win32com.client as win32

# FUNCION PARA EJECUTAR QUERY Y DEVOLVER EL DATAFRAME DE PANDAS
def execute_sql_query(connection_string, sql_query1):
    connection = pyodbc.connect(connection_string)
    cursor = connection.cursor()
    result = pd.read_sql_query(sql_query1,connection)
    connection.close()
    return result

# FUNCION PARA MANDAR EMAIL Y ADJUNTAR EXCEL
def send_email_with_attachment(to_address, subject, body, attachment_path):
    outlook = win32.Dispatch('Outlook.Application')
    mail = outlook.CreateItem(0)
    mail.To = to_address
    mail.Subject = subject
    mail.Body = body
    mail.Attachments.Add(attachment_path)
    mail.Send()

# CREDENCIALES BD
server = 'Localhost'
database = 'private'
username = 'private'
password = 'private'
connection_string = f'DRIVER={{SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}'

# QUERIES QUE EJECUTAN AMBOS SP GUARDADOS EN BD HISWEB PRODUCCION
sql_query1 = 'EXEC ListadoPacientesMalLlenados'
sql_query2 = 'EXEC TotalPacientesMalLlenados'

# EJECTUA QUERIES Y OBTIENE RESULTADO COMO UN DATAFRAME(SOLAMENTE EL sql_query1)
result_df = execute_sql_query(connection_string, sql_query1)
result2 = execute_sql_query(connection_string, sql_query2)
resultList = list(map(int, re.findall('\d+', str(result2))))
resultString = ' '.join(str (e) for e in resultList)
#for x in str(resultList):
#   resultString += ' ' + x
#result_a_INT = int(resultINT)
#resultFOR = [x [1] for x in result_a_INT]

# GUARDA RESULTADO COMO ARCHIVO EXCEL
excel_file_path = r"C:\Users\Administrador\Desktop\Pacientes mal llenados excel\HGZ98\PacientesMalLlenados_HGZ98.xlsx"
result_df.to_excel(excel_file_path, index=False)

# REMITENTE Y CUERPO DEL CORREO
yesterday = datetime.now() - timedelta(days=1)
week = datetime.now() - timedelta(days=7)
to_email = 'miguel.floreshe@imss.gob.mx; jcolvera@cmr3.com.mx'
email_subject = f"SMI DIG Incidencias presentadas entre fechas {week.strftime('%Y-%m-%d')} y {yesterday.strftime('%Y-%m-%d')}"
email_body = f"Buenos días Dr. Flores.\n\nEn la semana dentro de las fechas {week.strftime('%Y-%m-%d')} y {yesterday.strftime('%Y-%m-%d')}, en referencia al Servicio Médico Integral de Digitalización, Post-Procesamiento, Almacenamiento y Distribución de la Imagen en la unidad médica HGZ 98 le comento que NO hubo incidencias relacionadas con el servicio.\n\nCantidad de estudios mal llenados: {resultString[2:]} \n\nMantenimientos Correctivos: 0\n\nSaludos.\n\n*Correo automático"

# ENVIA CORREO CON EL EXCEL ADJUNTO
send_email_with_attachment(to_email, email_subject, email_body, excel_file_path)

print('CORREO ENVIADO EXITOSAMENTE.')




