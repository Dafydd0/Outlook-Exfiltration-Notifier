import os
import time
import win32com.client
import json
from datetime import datetime
import sys
from pathlib import PureWindowsPath, PurePosixPath

# Ruta del archivo de log
LOG_FILE = "C:\\Program Files (x86)\\ossec-agent\\active-response\\active-responses.log" if os.name == 'nt' else "/var/ossec/logs/active-responses.log"


def write_debug_file(ar_name, msg):
    with open(LOG_FILE, mode="a") as log_file:
        ar_name_posix = str(PurePosixPath(PureWindowsPath(ar_name[ar_name.find("active-response"):])))
        log_file.write(str(datetime.now().strftime('%Y/%m/%d %H:%M:%S')) + " " + ar_name_posix + ": " + msg + "\n")

# Redefinir la función print para incluir timestamp y escribir en el archivo de log
def print_with_timestamp(*args, **kwargs):
    """Print with a timestamp."""
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    message = ' '.join(map(str, args))
    write_debug_file(sys.argv[0], f"[{timestamp}] {message}")

# Mensaje inicial para verificar que la redirección funciona
print_with_timestamp("Script iniciado, redirección de logs establecida correctamente.")

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

# Función para asegurar que la carpeta 'Eventos' existe
def ensure_event_folder_exists(folder_name='Eventos'):
    if not os.path.exists(folder_name):
        os.makedirs(folder_name)

# Función para guardar en el archivo JSON dentro de la carpeta 'Eventos'
def save_to_json(data, folder_name='Eventos', filename='correos.json'):
    ensure_event_folder_exists(folder_name)
    file_path = os.path.join(folder_name, filename)
    
    try:
        with open(file_path, 'a') as f:
            for entry in data:
                json.dump(entry, f)
                f.write('\n')  # Agregar una línea nueva después de cada objeto JSON
        print_with_timestamp(f"Datos guardados exitosamente en {file_path}")
    except Exception as e:
        print_with_timestamp(f"Error al guardar datos en {file_path}: {e}")

# Función para cargar los mensajes existentes desde el archivo JSON
def load_existing_messages(folder_name='Eventos', filename='correos.json'):
    print_with_timestamp(f"Dentro de load_existing_messages()")
    file_path = os.path.join(folder_name, filename)
    if not os.path.exists(file_path):
        print_with_timestamp(f"Archivo no encontrado: {file_path}. Se retorna una lista vacía.")
        return []
    
    try:
        with open(file_path, 'r') as f:
            print_with_timestamp(f"Cargando mensajes desde {file_path}")
            return [json.loads(line) for line in f]
    except Exception as e:
        print_with_timestamp(f"Error al cargar mensajes desde {file_path}: {e}")
        return []

# Función para encontrar la carpeta de Enviados de la cuenta de Gmail dentro de Outlook
def get_sent_folder_for_gmail(namespace):
    try:
        print_with_timestamp(f"Dentro de get_sent_folder_for_gmail(namespace) antes del bucle")
        for account in namespace.Folders:
            print_with_timestamp(f"Dentro de get_sent_folder_for_gmail(namespace) dentro del bucle de cuentas")
            for folder in account.Folders:
                print_with_timestamp(f"Dentro de get_sent_folder_for_gmail(namespace) dentro del bucle de carpetas")
                if folder.Name == "[Gmail]":
                    print_with_timestamp(f"Dentro de get_sent_folder_for_gmail(namespace) dentro del bucle de cuentas (Gmail)")
                    for subfolder in folder.Folders:
                        print_with_timestamp(f"Dentro de get_sent_folder_for_gmail(namespace) dentro del bucle de subcuentas")
                        if subfolder.Name.lower() == "enviados":
                            print_with_timestamp(f"Dentro de get_sent_folder_for_gmail(namespace) dentro del bucle de subcuentas(enviados)")
                            return subfolder
    except Exception as e:
        print_with_timestamp(f"Error accessing folders for Gmail account: {e}")
    return None

# Función para obtener la carpeta de Enviados de Outlook
def get_sent_folder_for_outlook(namespace):
    try:
        print_with_timestamp(f"Dentro de get_sent_folder_for_outlook(namespace) dentro del try")
        sent = namespace.GetDefaultFolder(5)  # 5 representa la carpeta de enviados en Outlook
        print_with_timestamp(f"Dentro de get_sent_folder_for_outlook(namespace) despues de sent = namespace.GetDefaultFolder(5)")
        return sent
    except Exception as e:
        print_with_timestamp(f"Error accessing Sent folder: {e}")
        return None

# Función para determinar el tipo de cuenta y devolver la carpeta de Enviados correspondiente
def get_sent_folder():
    attempt_number = 0
    max_attempts = 5
    while attempt_number < max_attempts:
        try:
            print_with_timestamp(f"Dentro de get_sent_folder() dentro del try")
            time.sleep(5)
            print_with_timestamp(f"Dentro de get_sent_folder() 5s esperados antes de outlook...")
            outlook = win32com.client.Dispatch("Outlook.Application")
            print_with_timestamp(f"Dentro de get_sent_folder() después de outlook...")
            namespace = outlook.GetNamespace("MAPI")

            print_with_timestamp(f"namespace: "+ str(namespace))
            
            print_with_timestamp(f"Dentro de get_sent_folder() antes de llamar a get_sent_folder_for_outlook(namespace)")
            # Si no se encuentra la carpeta de Gmail, usar Outlook
            sent_folder = get_sent_folder_for_outlook(namespace)
            if sent_folder:
                print_with_timestamp("Using Outlook account.")
                return sent_folder
            
            print_with_timestamp(f"Dentro de get_sent_folder() antes de llamar a get_sent_folder_for_gmail(namespace)")
            # Intentar obtener la carpeta de Gmail
            sent_folder = get_sent_folder_for_gmail(namespace)
            if sent_folder:
                print_with_timestamp("Using Gmail account.")
                return sent_folder
                 

        except Exception as e:
            attempt_number += 1
            print_with_timestamp(f"Attempt {attempt_number}: Error initializing Outlook or accessing folders: {e}")
            if attempt_number >= max_attempts:
                print_with_timestamp("Max attempts reached. Exiting.")
                return None
            time.sleep(5)  # Esperar 5 segundos antes de reintentar

    return None

# Función para procesar nuevos mensajes y guardar solo los nuevos en el archivo JSON
def process_new_messages(sent_folder):
    try:
        print_with_timestamp(f"Dentro de process_new_messages() en el try")
        existing_messages = load_existing_messages()  # Cargar mensajes existentes desde el archivo JSON
        existing_subjects = {entry['Subject'] for entry in existing_messages}
        new_entries = []

        # Iterar sobre los mensajes no procesados en la carpeta de Enviados
        for message in sent_folder.Items:
            #print_with_timestamp(f"Dentro de process_new_messages() en el bucle")
            if message.Subject not in existing_subjects:
                print_with_timestamp(f"Dentro de process_new_messages() en el bucle en el if")
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
                print_with_timestamp(f"Dentro de process_new_messages() nuevo mensaje: {data}")
                print_with_timestamp(f"Dentro de process_new_messages() new_entries: {new_entries}")

        # Guardar la información solo si hay nuevos mensajes
        if new_entries:
            save_to_json(new_entries)

    except Exception as e:
        print_with_timestamp(f"Error processing emails: {e}")

def log_current_user():
    try:
        write_debug_file(sys.argv[0], "Dentro de log_current_user()")
        user = os.getlogin()  # Obtiene el usuario que ha iniciado sesión
        write_debug_file(sys.argv[0], f"Script ejecutado por el usuario: {user}")
    except Exception as e:
        write_debug_file(sys.argv[0], f"Error al obtener el usuario actual: {e}")
        

# Obtener la carpeta de Enviados
print_with_timestamp(f"Llamando a sent_folder = get_sent_folder()")
log_current_user()
sent_folder = get_sent_folder()


if not sent_folder:
    print_with_timestamp("Sent folder not found.")
    #exit()
else:
    print_with_timestamp(f"Total items in Sent folder: {sent_folder.Items.Count}")

# Bucle principal para verificar nuevos mensajes cada 10 segundos
while True:
    try:
        print_with_timestamp("Ejecutando Código")
        process_new_messages(sent_folder)
        time.sleep(5)  # Esperar 5 segundos antes de volver a verificar
    except Exception as e:
        print_with_timestamp(f"Error in main loop: {e}")
        time.sleep(5)
