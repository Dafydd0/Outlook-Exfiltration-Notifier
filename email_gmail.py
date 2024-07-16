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
    with open('correos_gmail.json', 'a') as f:
        for entry in data:
            json.dump(entry, f)
            f.write('\n')  # Agregar una línea nueva después de cada objeto JSON

# Función para cargar los mensajes existentes desde el archivo JSON
def load_existing_messages():
    existing_messages = []
    if os.path.exists('correos_gmail.json'):
        with open('correos_gmail.json', 'r') as f:
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

# Función para encontrar la carpeta de enviados de la cuenta de Gmail dentro de Outlook
def get_sent_folder_for_gmail():
    try:
        for account in namespace.Folders:
            if "gmail.com" in account.Name:  # Asegúrate de que este filtro se ajuste al nombre de tu cuenta
                for folder in account.Folders:
                    if folder.Name.lower() == "sent items" or folder.Name.lower() == "enviados":
                        return folder
    except Exception as e:
        print(f"Error accessing folders for Gmail account: {e}")
    return None

# Obtener la carpeta de enviados para la cuenta de Gmail dentro de Outlook
sent_folder = get_sent_folder_for_gmail()

if not sent_folder:
    print(f"Sent folder for Gmail account not found.")
    exit()
else:
    print(f"Total items in Sent folder for Gmail account: {sent_folder.Items.Count}")

# Función para procesar nuevos mensajes y guardar solo los nuevos en el archivo JSON
def process_new_messages():
    try:
        existing_messages = load_existing_messages()  # Cargar mensajes existentes desde el archivo JSON
        existing_subjects = {entry['Subject'] for entry in existing_messages}
        new_entries = []

        # Iterar sobre los mensajes no procesados en la carpeta de enviados
        for message in sent_folder.Items:
            if message.Subject not in existing_subjects:
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
        process_new_messages()
        time.sleep(10)  # Esperar 10 segundos antes de volver a verificar
    except Exception as e:
        print(f"Error in main loop: {e}")
        time.sleep(10)
