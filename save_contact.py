import win32com.client
import os
from datetime import datetime, timedelta
import csv

outlook = win32com.client.Dispatch("Outlook.Application")
mapi = outlook.GetNamespace('MAPI')

today = datetime.today()

#120 = intervalo de consulta en segundos
#segundos = (today.timestamp()-120)

#interval = datetime.fromtimestamp(segundos)

start_time = today.replace(day=1, hour=0, minute=0, second=0).strftime('%Y-%m-%d %H:%M %p')

#start_time = interval.strftime('%Y-%m-%d %H:%M %p')

end_time = today.strftime('%Y-%m-%d %H:%M %p')

messages = mapi.Folders("info@pulseaapp.com").Folders("Bandeja de entrada").Items
messages = messages.Restrict("[ReceivedTime] >= '" + start_time
+ "' And [ReceivedTime] <= '" + end_time + "'")
messages = messages.Restrict("[Subject] = 'Solicitud de información PULSEA'") 

contactos = []

for msg in list(messages):

    info = msg.Body.replace("Nombre: ","")
    info = info.replace("Numero: ","")
    info = info.replace("Correo: ","") 
    info = info.replace("Mensaje: ","")
    
    info_split = info.split("\r\n")
    json = {}
    json["Nombre"] = info_split[1]
    json["Numero"] = info_split[2]
    json["Correo"] = info_split[3]
    json["Mensaje"] = info_split[4]
    contactos.append(json)

for item in contactos:
    print(item)


with open("pulsea_contactos.csv","w",newline="") as file:
    csv_file = csv.writer(file)
    csv_file.writerow(["Nombre","Numero","Correo","Mensaje"])
    for item in contactos:
        csv_file.writerow([item["Nombre"],item["Numero"],item["Correo"],item["Mensaje"]])
