import pandas as pd
from pathlib import Path
import re

# Ruta a la carpeta de logs
carpeta_logs = Path(r"C:\Users\3002975\Documents\Py Proyect\leer_log_py")

# Lista para almacenar los tests
tests_extraidos = []

# Verificamos si la carpeta existe
if carpeta_logs.exists() and carpeta_logs.is_dir():
    for archivo_log in carpeta_logs.glob("*.txt"):
        with archivo_log.open("r", encoding="utf-8") as archivo:
            contenido = archivo.read()

            # Dividir por bloques de test
            bloques = re.split(r"(?=Test Description:)", contenido)

            for bloque in bloques:
                if "Test Description:" in bloque:
                    # Extraer campos usando regex o búsqueda directa
                    test_desc = re.search(r"Test Description:\s*(.*)", bloque)
                    test_num = re.search(r"Test (\d+)", bloque)
                    pcb_sn = re.search(r"PCB Serial Number:\s*(.*)", bloque)
                    lower = re.search(r"Test Lower Limit:\s*(.*)", bloque)
                    upper = re.search(r"Test Upper Limit:\s*(.*)", bloque)
                    measure = re.search(r"Test Measurement:\s*(.*)", bloque)
                    units = re.search(r"Units:\s*(.*)", bloque)
                    temp_start = re.search(r"Starting Temperature.*:\s*(.*)", bloque)
                    temp_end = re.search(r"Ending Temperature.*:\s*(.*)", bloque)
                    result = re.search(r"Test Result:\s*(.*)", bloque)

                    tests_extraidos.append({
                        "Archivo": archivo_log.name,
                        "Test Nº": test_num.group(1) if test_num else None,
                        "Descripción": test_desc.group(1).strip() if test_desc else None,
                        "Serie PCB": pcb_sn.group(1).strip() if pcb_sn else None,
                        "Límite Inferior": lower.group(1).strip() if lower else None,
                        "Límite Superior": upper.group(1).strip() if upper else None,
                        "Medida": measure.group(1).strip() if measure else None,
                        "Unidades": units.group(1).strip() if units else None,
                        "Temp Inicio": temp_start.group(1).strip() if temp_start else None,
                        "Temp Fin": temp_end.group(1).strip() if temp_end else None,
                        "Resultado": result.group(1).strip() if result else None
                    })

    # Crear el DataFrame
    df_tests = pd.DataFrame(tests_extraidos)

    # Guardar en Excel
    df_tests.to_excel("tests_extraidos.xlsx", index=False)
    print("Todos los tests han sido extraídos y guardados en 'tests_extraidos.xlsx'.")
else:
    print("Carpeta no encontrada:", carpeta_logs)
