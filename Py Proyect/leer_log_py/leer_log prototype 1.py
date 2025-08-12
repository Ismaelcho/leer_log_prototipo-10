from pathlib import Path

# Ruta completa al archivo de log
ruta_log = Path(r"C:\Users\3002975\Documents\Py Proyect\leer_log_py\ejemplo.log")

if ruta_log.exists():
    with ruta_log.open("r") as archivo:
        linea = archivo.read()  # Leer toda la línea (o todo el archivo)
        eventos = linea.split(";")  # Dividir por punto y coma

        print("Eventos con ERROR:\n")
        for evento in eventos:
            evento = evento.strip()
            if "ERROR" in evento:
                print(evento)
else:
    print("Archivo no encontrado:", ruta_log)

