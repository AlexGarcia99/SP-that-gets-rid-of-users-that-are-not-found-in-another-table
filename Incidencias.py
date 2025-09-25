#HGZ98
from datetime import datetime, timedelta
import re
import pyodbc  # para SQL Server
import win32com.client  # para Outlook automation

# Parametros de conexion de la BD
db_config = {
    'server': 'Localhost',
    'database': 'private',
    'user': 'private',
    'password': 'private',
}

# Configuracion de OutLook
yesterday = datetime.now() - timedelta(days=1)
outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")
recipient = "miguel.floreshe@imss.gob.mx"  # correo del jefe de servicio
subject = f"Estudios realizados el {yesterday.strftime('%Y-%m-%d')} HGZ98"
body = "Buenos días."

def execute_sql_query(query):
    connection = pyodbc.connect(
        f"DRIVER={{SQL Server}};SERVER={db_config['server']};DATABASE={db_config['database']};"
        f"UID={db_config['user']};PWD={db_config['password']}"
    )
    cursor = connection.cursor()
    cursor.execute(query)
    result = cursor.fetchall()
    connection.close()
    return result

def send_email(to_list, subject, body):
    mail = outlook.CreateItem(0)
    mail.To = ";".join(to_list)
    mail.Subject = subject
    mail.Body = body
    mail.Send()

def main():
    additional_recipient = "jcolvera@cmr3.com.mx"  #email extra a copiar en el correo
    # Query a la BD de MS SQL Server del total de estudios de un dia anteior(tomando en cuenta los VisiblePACS=1 y filtrando posibles UIDEstudio repetidos del dia). Apuntar a la BD(inmediatamente despues del FROM) en cuestion en el query.
    sql_query = "SELECT COUNT(W.PARTICION_UID) AS ESTUDIOS FROM (SELECT * FROM (SELECT ROW_NUMBER() OVER (PARTITION BY FOLIO ORDER BY (SELECT 1)) AS PARTICION_UID, * FROM [HIS_WEB].[dbo].[ImagenologiaEstudios] WHERE DATEADD(day, -1, convert(date, GETDATE())) = CONVERT(DATE, FECHAESTUDIO) AND VisiblePACS=1) AS T WHERE T.PARTICION_UID=1) AS W"

    # Ejecuta el query
    result = execute_sql_query(sql_query)
    resultINT = list(map(int, re.findall('\d+', str(result))))
    #resultSTRING = str(resultINT)
    yesterday = datetime.now() - timedelta(days=1)
    
    # Procesa el resultado del query en el cuerpo del correo
    email_body = f"Buenos días Dr. Flores.\n\nSe informa sobre la cantidad de estudios realizados.\n\n {yesterday.strftime('%Y-%m-%d')} : {resultINT} estudios. \n\nNota: Los estudios reportados pueden variar al final del mes debido a reconciliaciones solicitadas por el jefe de servicio o envío tardío de los mismos al servidor de imágenes.\n\nSaludos.\n\n*Correo automático" #concatenar el resultado del query(convertido a string) en esta linea
    #for row in result:
        #email_body += f"{row}\n"

    # Envia email
    send_email([recipient, additional_recipient], subject, email_body)

if __name__ == "__main__":
    main()

