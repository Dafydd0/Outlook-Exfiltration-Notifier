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
def save_to_json(data, filename='correos.json'):
    with open(filename, 'a') as f:
        for entry in data:
            json.dump(entry, f)
            f.write('\n')  # Agregar una línea nueva después de cada objeto JSON

# Función para cargar los mensajes existentes desde el archivo JSON
def load_existing_messages(filename='correos.json'):
    if not os.path.exists(filename):
        return []
    
    with open(filename, 'r') as f:
        return [json.loads(line) for line in f]

# Función para encontrar la carpeta de Enviados de la cuenta de Gmail dentro de Outlook
def get_sent_folder_for_gmail(namespace):
    try:
        for account in namespace.Folders:
            for folder in account.Folders:
                if folder.Name == "[Gmail]":
                    for subfolder in folder.Folders:
                        if subfolder.Name.lower() == "enviados":
                            return subfolder
    except Exception as e:
        print(f"Error accessing folders for Gmail account: {e}")
    return None

# Función para obtener la carpeta de Enviados de Outlook
def get_sent_folder_for_outlook(namespace):
    try:
        sent = namespace.GetDefaultFolder(5)  # 5 representa la carpeta de enviados en Outlook
        return sent
    except Exception as e:
        print(f"Error accessing Sent folder: {e}")
        return None

# Función para determinar el tipo de cuenta y devolver la carpeta de Enviados correspondiente
def get_sent_folder():
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        namespace.Logon()  # Esto debería pedir el perfil si está configurado así

        # Intentar obtener la carpeta de Gmail
        sent_folder = get_sent_folder_for_gmail(namespace)
        if sent_folder:
            print("Using Gmail account.")
            return sent_folder

        # Si no se encuentra la carpeta de Gmail, usar Outlook
        sent_folder = get_sent_folder_for_outlook(namespace)
        if sent_folder:
            print("Using Outlook account.")
            return sent_folder

    except Exception as e:
        print(f"Error initializing Outlook or accessing folders: {e}")
        return None

# Función para procesar nuevos mensajes y guardar solo los nuevos en el archivo JSON
def process_new_messages(sent_folder):
    try:
        existing_messages = load_existing_messages()  # Cargar mensajes existentes desde el archivo JSON
        existing_subjects = {entry['Subject'] for entry in existing_messages}
        new_entries = []

        # Iterar sobre los mensajes no procesados en la carpeta de Enviados
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

# Obtener la carpeta de Enviados
sent_folder = get_sent_folder()

if not sent_folder:
    print(f"Sent folder not found.")
    exit()
else:
    print(f"Total items in Sent folder: {sent_folder.Items.Count}")

# Bucle principal para verificar nuevos mensajes cada 10 segundos
while True:
    try:
        process_new_messages(sent_folder)
        time.sleep(10)  # Esperar 10 segundos antes de volver a verificar
    except Exception as e:
        print(f"Error in main loop: {e}")
        time.sleep(10)
