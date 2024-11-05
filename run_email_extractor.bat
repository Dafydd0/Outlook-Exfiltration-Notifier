@echo off
echo %date% %time% - run_email_extractor.bat executed >> "C:\Program Files (x86)\ossec-agent\active-response\bin\execution.log"
cd "C:\Program Files (x86)\ossec-agent\active-response\bin"

:: Cambia la ruta a donde esté instalado Python
"C:\Users\David\AppData\Local\Programs\Python\Python312\python.exe" email_extractor.py >> "C:\Program Files (x86)\ossec-agent\active-response\bin\execution.log" 2>&1

echo Ejecución completada >> "C:\Program Files (x86)\ossec-agent\active-response\bin\execution.log"
