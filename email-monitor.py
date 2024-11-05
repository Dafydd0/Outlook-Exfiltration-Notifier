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
        
            write_debug_file(argv[0], 'Antes de llamar al extractor de correos')
            run_email_extractor()


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

def run_email_extractor():
    try:
        write_debug_file(sys.argv[0], "Ejecutando .bat")
        subprocess.run(['C:\\Program Files (x86)\\ossec-agent\\active-response\\bin\\run_email_extractor.bat'], check=True)
        write_debug_file(sys.argv[0], ".bat Ejecutado")
        
    except subprocess.CalledProcessError as e:
        write_debug_file(f"Error al ejecutar el archivo .bat: {e}")

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
