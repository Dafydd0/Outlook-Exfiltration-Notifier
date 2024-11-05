#!/usr/bin/python3
# Copyright (C) 2015-2022, Wazuh Inc.
# All rights reserved.

# This program is free software; you can redistribute it
# and/or modify it under the terms of the GNU General Public
# License (version 2) as published by the FSF - Free Software
# Foundation.

import os
import sys
import json
import datetime
import time
import threading
from pathlib import PureWindowsPath, PurePosixPath
import win32com.client
import subprocess

if os.name == 'nt':
    # Obtener el directorio desde el cual se está ejecutando el script
    current_dir = os.path.dirname(os.path.abspath(__file__))

    # Definir la ruta del archivo de log en el directorio actual
    #LOG_FILE = os.path.join('C:\\Users\\David\\Desktop\\MEGA\\Trabajo\\Pruebas\\Email', 'active-responses.log')

    LOG_FILE = "C:\\Program Files (x86)\\ossec-agent\\active-response\\active-responses.log"

else:
    LOG_FILE = "/var/ossec/logs/active-responses.log"

ADD_COMMAND = 0
DELETE_COMMAND = 1
CONTINUE_COMMAND = 2
ABORT_COMMAND = 3

OS_SUCCESS = 0
OS_INVALID = -1


class Message:
    def __init__(self):
        self.alert = ""
        self.command = 0

# Función para escribir en el archivo de depuración
def write_debug_file(ar_name, msg):
    with open(LOG_FILE, mode="a") as log_file:
        ar_name_posix = str(PurePosixPath(PureWindowsPath(ar_name[ar_name.find("active-response"):])))
        log_file.write(str(datetime.datetime.now().strftime('%Y/%m/%d %H:%M:%S')) + " " + ar_name_posix + ": " + msg + "\n")

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
    
    with open(file_path, 'a') as f:
        for entry in data:
            json.dump(entry, f)
            f.write('\n')  # Agregar una línea nueva después de cada objeto JSON

# Función para cargar los mensajes existentes desde el archivo JSON
def load_existing_messages(folder_name='Eventos', filename='correos.json'):
    file_path = os.path.join(folder_name, filename)
    if not os.path.exists(file_path):
        return []
    
    with open(file_path, 'r') as f:
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
    write_debug_file(sys.argv[0], "Dentro de get_sent_folder_for_outlook()")
    try:
        write_debug_file(sys.argv[0], "Dentro de get_sent_folder_for_outlook() antes de hacer sent = namespace.GetDefaultFolder(5)")
        attempt_number = 0
        max_attempts = 5
        while attempt_number < max_attempts:
            try:
                sent = namespace.GetDefaultFolder(5)
                write_debug_file(sys.argv[0], "Dentro de get_sent_folder_for_outlook() sent folder obtenida")
                return sent
            except Exception as e:
                attempt_number += 1
                write_debug_file(sys.argv[0], f"Attempt {attempt_number}: Error accessing Sent folder: {e}")
                time.sleep(2)  # Esperar antes de intentar de nuevo
    except win32com.client.pywintypes.com_error as com_error:
        write_debug_file(sys.argv[0], f"COM error accessing Sent folder: {com_error}")
    except Exception as e:
        write_debug_file(sys.argv[0], f"Error accessing Sent folder: {e}")
    return None


# Función para determinar el tipo de cuenta y devolver la carpeta de Enviados correspondiente
def get_sent_folder():
    try:
        write_debug_file(sys.argv[0], "Dentro de get_sent_folder() antes de hacer el login de Outlook")
        outlook = win32com.client.Dispatch("Outlook.Application")
        write_debug_file(sys.argv[0], "Dentro de get_sent_folder() despues de outlook = win32com.client.Dispatch('Outlook.Application')")
        namespace = outlook.GetNamespace("MAPI")
        #namespace.Logon()  # Esto debería pedir el perfil si está configurado así
        write_debug_file(sys.argv[0], "Dentro de get_sent_folder() depués de hacer el login de Outlook")
        # Intentar obtener la carpeta de Gmail
        #sent_folder = get_sent_folder_for_gmail(namespace)
       # if sent_folder:
        #    print("Using Gmail account.")
        #    return sent_folder

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

# Función para configurar y verificar el mensaje
def setup_and_check_message(argv):
    # get alert from stdin
    input_str = ""
    for line in sys.stdin:
        input_str = line
        break

    write_debug_file(argv[0], input_str)

    try:
        data = json.loads(input_str)
    except ValueError:
        write_debug_file(argv[0], 'Decoding JSON has failed, invalid input format')
        msg = Message()
        msg.command = OS_INVALID
        return msg

    msg = Message()
    msg.alert = data

    command = data.get("command")

    if command == "add":
        msg.command = ADD_COMMAND
    elif command == "delete":
        msg.command = DELETE_COMMAND
    else:
        msg.command = OS_INVALID
        write_debug_file(argv[0], 'Not valid command: ' + command)

    return msg

# Función para enviar claves y verificar el mensaje
def send_keys_and_check_message(argv, keys):
    # build and send message with keys
    keys_msg = json.dumps({"version": 1,"origin":{"name": argv[0],"module":"active-response"},"command":"usb-monitor-autoclose","parameters":{"keys":keys}})

    write_debug_file(argv[0], keys_msg)

    print(keys_msg)
    sys.stdout.flush()

    # read the response of previous message
    input_str = ""
    while True:
        line = sys.stdin.readline()
        if line:
            input_str = line
            break

    write_debug_file(argv[0], input_str)

    try:
        data = json.loads(input_str)
    except ValueError:
        write_debug_file(argv[0], 'Decoding JSON has failed, invalid input format')
        return Message

    action = data.get("command")

    if "continue" == action:
        ret = CONTINUE_COMMAND
    elif "abort" == action:
        ret = ABORT_COMMAND
    else:
        ret = OS_INVALID
        write_debug_file(argv[0], "Invalid value of 'command'")

    return ret

# Función principal del script
def main(argv):
    try:
        write_debug_file(argv[0], "Llamando a log_current_user()")
        log_current_user()
        write_debug_file(argv[0], "Started")
        # Validar json y obtener el comando
        msg = setup_and_check_message(argv)

        if msg.command < 0:
            sys.exit(OS_INVALID)
        if msg.command == ADD_COMMAND:
            alert = msg.alert["parameters"]["alert"]
            keys = [alert["rule"]["id"]]

            action = send_keys_and_check_message(argv, keys)

            if action != CONTINUE_COMMAND:
                if action == ABORT_COMMAND:
                    write_debug_file(argv[0], "Aborted")
                    sys.exit(OS_SUCCESS)
                else:
                    write_debug_file(argv[0], "Invalid command")
                    sys.exit(OS_INVALID)

            # Ejecutar monitor de correos en un hilo separado
            
            write_debug_file(argv[0], 'Antes de llamar al monitor de correos')
            #write_debug_file(argv[0], 'argv[0]' + str(argv[0]))
            #email_monitor_thread = threading.Thread(target=email_monitor)
            #email_monitor_thread.start()

            #email_monitor()

            # Esperar a que los hilos terminen
            #email_monitor_thread.join()
            #email_extractor_thread.join()

            # Ejecutar el script email_extractor.exe en otro hilo
            write_debug_file(argv[0], 'Antes de llamar al extractor de correos')
            run_email_extractor()
            #email_extractor_thread = threading.Thread(target=run_email_extractor)
            #email_extractor_thread.start()

            write_debug_file(argv[0], 'Después de ejecutar hilos')

        elif msg.command == DELETE_COMMAND:
            # Similar a lo anterior, pero para el comando DELETE_COMMAND
            pass

    except Exception as e:
        write_error_file(argv[0], str(e))
        raise  # Esto volverá a lanzar la excepción para depuración si es necesario.

def write_error_file(ar_name, msg):
    with open("errors.log", mode="a") as log_file:
        ar_name_posix = str(PurePosixPath(PureWindowsPath(ar_name[ar_name.find("active-response"):])))
        log_file.write(str(datetime.datetime.now().strftime('%Y/%m/%d %H:%M:%S')) + " " + ar_name_posix + ": " + msg + "\n")


# Función para el monitoreo de correos
def email_monitor():

    write_debug_file(sys.argv[0], "Dentro de email_monitor() pero sin llamar a get_sent_folder()")
    sent_folder = get_sent_folder()
    write_debug_file(sys.argv[0], "Después de llamar a get_sent_folder()")

    if not sent_folder:
        write_debug_file(sys.argv[0], "Sent folder not found.")
        return
    
    write_debug_file(sys.argv[0], f"Total items in Sent folder: {sent_folder.Items.Count}")

    # Bucle principal para verificar nuevos mensajes cada 10 segundos
    while True:
        try:
            process_new_messages(sent_folder)
            time.sleep(10)  # Esperar 10 segundos antes de volver a verificar
        except Exception as e:
            write_debug_file(sys.argv[0], f"Error in main loop: {e}")
            time.sleep(10)  # Espera antes de intentar de nuevo en caso de error

# Función para ejecutar el script email_extractor.exe
import subprocess

'''
def run_email_extractor():
    try:
        # Nombre de la tarea programada
        task_name = "RunEmailExtractorAsDavid"
        email_extractor_path = r'C:\"Program Files (x86)"\ossec-agent\active-response\bin\email_extractor.exe'
        
        # Establecer una hora en el pasado cercano para evitar advertencias
        now = datetime.datetime.now()
        start_time = (now + datetime.timedelta(seconds=5)).strftime('%H:%M')

        # Configurar la tarea programada si no existe
        command_create = (
            f'schtasks /create /tn "{task_name}" /tr "{email_extractor_path}" /sc once '
            f'/st {start_time} /ru David /rl highest /f'
        )
        result_create = subprocess.run(command_create, shell=True, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        
        write_debug_file(sys.argv[0], f"Tarea programada creada o actualizada. stdout: {result_create.stdout}")
        write_debug_file(sys.argv[0], f"Tarea programada creada o actualizada. stderr: {result_create.stderr}")

        # Ejecutar la tarea programada con schtasks
        command_run = f'schtasks /run /tn "{task_name}"'
        result_run = subprocess.run(command_run, shell=True, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        
        write_debug_file(sys.argv[0], f"Script extractor de emails ejecutado. stdout: {result_run.stdout}")
        write_debug_file(sys.argv[0], f"Script extractor de emails ejecutado. stderr: {result_run.stderr}")
        write_debug_file(sys.argv[0], f"Script extractor de emails finalizado con código {result_run.returncode}")

    except subprocess.CalledProcessError as cpe:
        write_debug_file(sys.argv[0], f"Error al ejecutar la tarea con código de retorno: {cpe.returncode}")
        write_debug_file(sys.argv[0], f"stdout: {cpe.stdout}")
        write_debug_file(sys.argv[0], f"stderr: {cpe.stderr}")
    except Exception as e:
        write_debug_file(sys.argv[0], f"Error al configurar o ejecutar la tarea programada: {e}")
'''


def run_email_extractor():
    try:
        write_debug_file(sys.argv[0], "Ejecutando .bat")
        subprocess.run(['C:\\Program Files (x86)\\ossec-agent\\active-response\\bin\\run_email_extractor.bat'], check=True)

    except subprocess.CalledProcessError as cpe:
        write_debug_file(sys.argv[0], f"Error al ejecutar el script con código de retorno: {cpe.returncode}")
        write_debug_file(sys.argv[0], f"stdout: {cpe.stdout}")
        write_debug_file(sys.argv[0], f"stderr: {cpe.stderr}")
    except Exception as e:
        write_debug_file(sys.argv[0], f"Error al ejecutar el archivo .bat: {e}")

# Función para obtener y escribir el nombre del usuario actual en el archivo de log
def log_current_user():
    try:
        write_debug_file(sys.argv[0], "Dentro de log_current_user()")
        user = os.getlogin()  # Obtiene el usuario que ha iniciado sesión
        write_debug_file(sys.argv[0], f"Script ejecutado por el usuario: {user}")
    except Exception as e:
        write_debug_file(sys.argv[0], f"Error al obtener el usuario actual: {e}")



if __name__ == "__main__":
    main(sys.argv)
