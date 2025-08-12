import pandas as pd
from pathlib import Path
import re

# Ruta a la carpeta que contiene los archivos .log
carpeta_logs = Path(r"C:\Users\3002975\Documents\Py Proyect\leer_log_py")

# Expresión regular para extraer fecha, tipo, descripción y ubicación
patron = re.compile(r"\[(.*?)\]\s+(ERROR):\s+(.*?)\s+en\s+(.*)")

# Lista para almacenar todos los errores encontrados
errores = []

# Verificar si la carpeta existe
if carpeta_logs.exists() and carpeta_logs.is_dir():
    # Iterar sobre todos los archivos .log en la carpeta
    for archivo_log in carpeta_logs.glob("*.txt"):
        with archivo_log.open("r", encoding="utf-8") as archivo:
            contenido = archivo.read()
            eventos = contenido.split(";")
            
            for evento in eventos:
                evento = evento.strip()
                if "ERROR" in evento:
                    coincidencia = patron.search(evento)
                    if coincidencia:
                        fecha, tipo, descripcion, ubicacion = coincidencia.groups()
                        errores.append({
                            "Archivo": archivo_log.name,
                            "Fecha": fecha,
                            "Tipo": tipo,
                            "Descripción": descripcion,
                            "Ubicación": ubicacion
                        })

    # Crear un DataFrame y guardarlo en Excel
    df = pd.DataFrame(errores)
    df.to_excel("errores_log_completo.xlsx", index=False)
    print("Los eventos con ERROR de todos los archivos han sido guardados en 'errores_log_completo.xlsx'.")
else:
    print("Carpeta no encontrada:", carpeta_logs)

