import os
import time
import win32com.client
import json
from datetime import datetime

# Función para obtener la dirección de correo electrónico del remitente
def get_sender_email_address(mail):
    sender = mail.Sender
    sender_email_address = ""

    if sender.AddressEntryUserType == 0 or sender.AddressEntryUserType == 5:
        exch_user = sender.GetExchangeUser()
        if exch_user is not None:
            sender_email_address = exch_user.PrimarySmtpAddress
    else:
        sender_email_address = mail.SenderEmailAddress

    return sender_email_address

# Función para obtener el nombre de usuario y el nombre del PC
def get_user_and_pc():
    user = os.getlogin()
    pc = os.getenv('COMPUTERNAME')
    return user, pc

# Función para guardar en el archivo JSON
def save_to_json(data):
    # Escribir todos los datos en el archivo JSON
    with open('correos.json', 'a') as f:
        for entry in data:
            json.dump(entry, f)
            f.write('\n')  # Agregar una línea nueva después de cada objeto JSON

# Función para cargar los mensajes existentes desde el archivo JSON
def load_existing_messages():
    existing_messages = []
    if os.path.exists('correos.json'):
        with open('correos.json', 'r') as f:
            for line in f:
                line = line.strip()
                if line:
                    try:
                        existing_messages.append(json.loads(line))
                    except json.JSONDecodeError as e:
                        print(f"Error loading JSON line: {e}")
    return existing_messages

# Iniciar la aplicación de Outlook
try:
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    namespace.Logon()  # Esto debería pedir el perfil si está configurado así
except Exception as e:
    print(f"Error initializing Outlook: {e}")
    exit()

# Comprobar las carpetas disponibles
try:
    for i in range(20):  # Límite arbitrario para verificar la existencia de carpetas
        try:
            folder = namespace.GetDefaultFolder(i)
            print(f"Folder {i}: {folder.Name}")
        except Exception as e:
            print(f"Folder {i} not accessible: {e}")
except Exception as e:
    print(f"Error accessing folders: {e}")
    exit()

# Obtener la carpeta de enviados
try:
    sent = namespace.GetDefaultFolder(5)  # 5 representa la carpeta de enviados en Outlook
    print(f"Total items in Sent folder: {sent.Items.Count}")
except Exception as e:
    print(f"Error accessing Sent folder: {e}")
    exit()

# Función para procesar nuevos mensajes y guardar solo los nuevos en el archivo JSON
def process_new_messages():
    try:
        existing_messages = load_existing_messages()  # Cargar mensajes existentes desde el archivo JSON
        new_entries = []

        # Iterar sobre los mensajes no procesados en la carpeta de enviados
        unreprocessed_messages = [message for message in sent.Items if message.Subject not in [entry['Subject'] for entry in existing_messages]]

        for message in unreprocessed_messages:
            sender = get_sender_email_address(message)
            user, pc = get_user_and_pc()

            # Construir el objeto JSON para cada mensaje
            data = {
                "Date": message.CreationTime.strftime('%Y-%m-%d %H:%M:%S'),
                "User": user,
                "PC": pc,
                "To": message.To,
                "CC": message.CC,
                "BCC": message.BCC,
                "From": sender,
                "Size": message.Size,
                "Attachments": [attachment.FileName for attachment in message.Attachments],
                "Content": message.Body,
                "Subject": message.Subject
            }
            new_entries.append(data)

        # Guardar la información solo si hay nuevos mensajes
        if new_entries:
            save_to_json(new_entries)

    except Exception as e:
        print(f"Error processing emails: {e}")

# Bucle principal para verificar nuevos mensajes cada 10 segundos
while True:
    try:
        # Procesar nuevos mensajes
        process_new_messages()

        # Esperar 10 segundos antes de volver a verificar
        time.sleep(10)

    except Exception as e:
        print(f"Error in main loop: {e}")
        time.sleep(10)