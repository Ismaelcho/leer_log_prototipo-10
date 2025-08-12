import pandas as pd
from pathlib import Path
import re

# Clase para representar un solo test
class LogTest:
    def __init__(self, archivo, test_num, descripcion, serie_pcb, limite_inf,
                 limite_sup, medida, unidades, temp_ini, temp_fin, resultado):
        self.archivo = archivo
        self.test_num = int(test_num) if test_num and test_num.isdigit() else None
        self.descripcion = descripcion
        self.serie_pcb = serie_pcb
        self.limite_inf = limite_inf
        self.limite_sup = limite_sup
        self.medida = medida
        self.unidades = unidades
        self.temp_ini = temp_ini
        self.temp_fin = temp_fin
        self.resultado = resultado

# Clase para analizar un archivo de log
class LogParser:
    def __init__(self, path_archivo):
        self.path_archivo = path_archivo

    def parse(self):
        tests = []
        with self.path_archivo.open("r", encoding="utf-8") as archivo:
            contenido = archivo.read()

            bloques = re.split(r"(?=Test Description:)", contenido)

            for bloque in bloques:
                if "Test Description:" not in bloque:
                    continue

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

                test = LogTest(
                    archivo=self.path_archivo.name,
                    test_num=test_num.group(1) if test_num else None,
                    descripcion=test_desc.group(1).strip() if test_desc else None,
                    serie_pcb=pcb_sn.group(1).strip() if pcb_sn else None,
                    limite_inf=lower.group(1).strip() if lower else None,
                    limite_sup=upper.group(1).strip() if upper else None,
                    medida=measure.group(1).strip() if measure else None,
                    unidades=units.group(1).strip() if units else None,
                    temp_ini=temp_start.group(1).strip() if temp_start else None,
                    temp_fin=temp_end.group(1).strip() if temp_end else None,
                    resultado=result.group(1).strip() if result else None,
                )

                tests.append(test)

        return tests

# Clase para procesar múltiples archivos y exportar resultados
class LogProcessor:
    def __init__(self, carpeta_logs):
        self.carpeta_logs = Path(carpeta_logs)
        self.tests = []

    def procesar_logs(self):
        if not self.carpeta_logs.exists() or not self.carpeta_logs.is_dir():
            print("❌ Carpeta no encontrada:", self.carpeta_logs)
            return

        for archivo_log in self.carpeta_logs.glob("*.txt"):
            parser = LogParser(archivo_log)
            self.tests.extend(parser.parse())

    def exportar_excel_personalizado(self, nombre_archivo="tests_formato_personalizado.xlsx"):
        # Agrupar por PCB y ordenar por número de test
        pcb_dict = {}
        for test in self.tests:
            pcb_dict.setdefault(test.serie_pcb, []).append(test)

        # Construir filas para el DataFrame
        filas = []
        for pcb, tests in pcb_dict.items():
            tests_ordenados = sorted(tests, key=lambda x: x.test_num if x.test_num is not None else 0)
            for test in tests_ordenados:
                filas.append({
                    "Test": test.descripcion,
                    "LowLimit": test.limite_inf,
                    "HighLimit": test.limite_sup,
                    "Measurement": test.medida,
                    "Trial1": test.medida,
                    "Trial2": "",
                    "Trial3": ""
                })
            filas.append({})  # línea vacía entre PCBs

        df = pd.DataFrame(filas)
        df.to_excel(nombre_archivo, index=False)
        print(f"✅ Archivo Excel exportado como '{nombre_archivo}' con pruebas ordenadas por número.")

# Función principal
def main():
    ruta = r"C:\Users\3002975\Documents\Py Proyect\leer_log_py"  # Ajustar si es necesario
    procesador = LogProcessor(ruta)
    procesador.procesar_logs()
    procesador.exportar_excel_personalizado()

if __name__ == "__main__":
    main()

