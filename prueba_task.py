import subprocess
import os

# Ruta del archivo de comandos
cmd_file_path = r'C:\Users\David\Desktop\MEGA\Trabajo\Pruebas\Email\run_email_extractor.cmd'

# Nombre de la tarea programada
task_name = "RunEmailExtractorAsDavid"

# Ruta del archivo de registro
log_file_path = 'prueba_task.log'

# Función para escribir mensajes en el archivo de registro
def log_message(message):
    with open(log_file_path, 'a') as log_file:
        log_file.write(message + '\n')

# Comando para crear la tarea programada
command_create = (
    f'schtasks /create /tn "{task_name}" /tr "{cmd_file_path}" /sc once '
    f'/st 00:00 /ru David /rl highest /f'
)

# Ejecutar el comando para crear la tarea
try:
    result_create = subprocess.run(
        command_create,
        shell=True,
        check=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True
    )
    log_message("Tarea programada creada exitosamente.")
    log_message(result_create.stdout)
except subprocess.CalledProcessError as e:
    log_message("Error al crear la tarea programada.")
    log_message(e.stderr)

# Comando para ejecutar la tarea programada inmediatamente
command_run = f'schtasks /run /tn "{task_name}"'

# Ejecutar el comando para ejecutar la tarea
try:
    result_run = subprocess.run(
        command_run,
        shell=True,
        check=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True
    )
    log_message("Tarea programada ejecutada.")
    log_message(result_run.stdout)
except subprocess.CalledProcessError as e:
    log_message("Error al ejecutar la tarea programada.")
    log_message(e.stderr)

# Verificar el estado de la tarea después de ejecutar
status_command = f'schtasks /query /tn "{task_name}" /v'
try:
    status_result = subprocess.run(
        status_command,
        shell=True,
        check=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True
    )
    log_message("Estado de la tarea programada:")
    log_message(status_result.stdout)
except subprocess.CalledProcessError as e:
    log_message("Error al consultar el estado de la tarea programada.")
    log_message(e.stderr)
