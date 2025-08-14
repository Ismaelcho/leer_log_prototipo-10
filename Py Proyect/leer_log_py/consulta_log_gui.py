import pandas as pd
from pathlib import Path
import re

# Clase para representar un solo test
class LogTest:
    def __init__(self, archivo, descripcion, serie_pcb, limite_inf,
                 limite_sup, medida, unidades, temp_ini, temp_fin, resultado):
        self.archivo = archivo
        self.descripcion = descripcion
        self.serie_pcb = serie_pcb
        self.limite_inf = limite_inf
        self.limite_sup = limite_sup
        self.medida = medida
        self.unidades = unidades
        self.temp_ini = temp_ini
        self.temp_fin = temp_fin
        self.resultado = resultado

# Función para limpiar la descripción eliminando "Test X"
def limpiar_descripcion(texto):
    return re.sub(r"Test\s*\d+\s*", "", texto).strip()

# Función para procesar un archivo log y extraer los datos
def procesar_log(ruta_archivo):
    tests = []
    with open(ruta_archivo, "r", encoding="utf-8") as file:
        for linea in file:
            match = re.search(
                r"(Test\s*\d+\s*[\w\s]+)\s+PCB Serial Number:\s*([\w-]+)\s+Limits:\s*([\d.-]+)\s*to\s*([\d.-]+)\s+Measured:\s*([\d.-]+)\s*(\w+)\s+Temp Start:\s*([\d.-]+)\s*C\s+Temp End:\s*([\d.-]+)\s*C\s+Result:\s*(\w+)",
                linea
            )
            if match:
                descripcion_limpia = limpiar_descripcion(match.group(1))
                tests.append(LogTest(
                    archivo=ruta_archivo.name,
                    descripcion=descripcion_limpia,
                    serie_pcb=match.group(2),
                    limite_inf=float(match.group(3)),
                    limite_sup=float(match.group(4)),
                    medida=float(match.group(5)),
                    unidades=match.group(6),
                    temp_ini=float(match.group(7)),
                    temp_fin=float(match.group(8)),
                    resultado=match.group(9)
                ))
    return tests

# Carpeta con los logs
carpeta_logs = Path("logs")
todos_los_tests = []

for archivo in carpeta_logs.glob("*.log"):
    todos_los_tests.extend(procesar_log(archivo))

# Crear DataFrame
df = pd.DataFrame([vars(test) for test in todos_los_tests])

# Guardar en Excel
df.to_excel("resultados_limpios.xlsx", index=False)

print("Archivo Excel generado correctamente sin 'Test X' en la descripción.")
