#!/usr/bin/python3
import os
import sys
import datetime
import subprocess

# Configurar el archivo de log dependiendo del sistema operativo
LOG_FILE = "C:\\Program Files (x86)\\ossec-agent\\active-response\\active-responses.log" if os.name == 'nt' else "/var/ossec/logs/active-responses.log"

# Función para escribir en el archivo de log
def write_log(msg):
    with open(LOG_FILE, mode="a") as log_file:
        log_file.write(f"{datetime.datetime.now().strftime('%Y/%m/%d %H:%M:%S')}: {msg}\n")

# Función para ejecutar el archivo .bat
def run_bat_file():
    try:
        write_log("Se va a jecutar el archivo .bat.")
        subprocess.run(['C:\\Program Files (x86)\\ossec-agent\\active-response\\bin\\run_email_extractor.bat'], check=True)
        write_log("El archivo .bat se ejecutó correctamente.")
    except subprocess.CalledProcessError as e:
        write_log(f"Error al ejecutar el archivo .bat: {e}")

# Función principal
def main():
    write_log("Iniciando el script.")
    run_bat_file()

if __name__ == "__main__":
    main()
