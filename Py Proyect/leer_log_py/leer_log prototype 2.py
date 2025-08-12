import pandas as pd
from pathlib import Path
import re

# Ruta al archivo de log
ruta_log = Path(r"C:\Users\3002975\Documents\Py Proyect\leer_log_py\ejemplo.log")

# Lista para almacenar los eventos con ERROR
errores = []

# Expresión regular para extraer fecha, tipo, descripción y ubicación
patron = re.compile(r"\[(.*?)\]\s+(ERROR):\s+(.*?)\s+en\s+(.*)")

if ruta_log.exists():
    with ruta_log.open("r") as archivo:
        contenido = archivo.read()
        eventos = contenido.split(";")
        
#Recorre cada evento, Elimina espacios al inicio y final con strip(),
#Solo procesa eventos que contengan la palabra ERROR,Aplica la expresión regular para extraer sus componentes.
        for evento in eventos:
            evento = evento.strip()
            if "ERROR" in evento:
                coincidencia = patron.search(evento)
                if coincidencia:
                    fecha, tipo, descripcion, ubicacion = coincidencia.groups()
                    errores.append({
                        "Fecha": fecha,
                        "Tipo": tipo,
                        "Descripción": descripcion,
                        "Ubicación": ubicacion
                    })

    # Crear un DataFrame y guardarlo en Excel
    df = pd.DataFrame(errores)
    df.to_excel("errores_log.xlsx", index=False)

#Si todo sale bien, muestra mensaje de éxito.
#Si el archivo .log no existe, lo avisa con un mensaje de error.
    print("Los eventos con ERROR han sido guardados en 'errores_log.xlsx'.")
else:
    print("Archivo no encontrado:", ruta_log)