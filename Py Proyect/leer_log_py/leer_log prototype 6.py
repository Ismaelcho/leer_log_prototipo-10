import pandas as pd
from pathlib import Path
import re

#1. LogTest (una clase para representar un solo test)
class LogTest:
    def __init__(self, archivo, test_num, descripcion, serie_pcb, limite_inf,
                 limite_sup, medida, unidades, temp_ini, temp_fin, resultado):
        self.archivo = archivo
        self.test_num = test_num
        self.descripcion = descripcion
        self.serie_pcb = serie_pcb
        self.limite_inf = limite_inf
        self.limite_sup = limite_sup
        self.medida = medida
        self.unidades = unidades
        self.temp_ini = temp_ini
        self.temp_fin = temp_fin
        self.resultado = resultado

    def to_dict(self):
        return {
            "Archivo": self.archivo,
            "Test Nº": self.test_num,
            "Descripción": self.descripcion,
            "Serie PCB": self.serie_pcb,
            "Límite Inferior": self.limite_inf,
            "Límite Superior": self.limite_sup,
            "Medida": self.medida,
            "Unidades": self.unidades,
            "Temp Inicio": self.temp_ini,
            "Temp Fin": self.temp_fin,
            "Resultado": self.resultado
        }

#2. LogParser (una clase para analizar un archivo de log)
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

#3. LogProcessor (para manejar múltiples archivos y guardar los resultados)
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

    def exportar_excel(self, nombre_archivo="tests_extraidos.xlsx"):
        df = pd.DataFrame([test.to_dict() for test in self.tests])
        df.to_excel(nombre_archivo, index=False)
        print(f"✅ Se han exportado {len(self.tests)} tests a '{nombre_archivo}'.")

#4. main() (función de entrada para ejecutar todo el flujo)
def main():
    ruta = r"C:\Users\3002975\Documents\Py Proyect\leer_log_py"  # Cambia esta ruta si es necesario
    procesador = LogProcessor(ruta)
    procesador.procesar_logs()
    procesador.exportar_excel()


if __name__ == "__main__":
    main()
