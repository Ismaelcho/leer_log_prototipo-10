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

# Clase para analiza un archivo de log
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

                # Verifica si alguno de los campos contiene 'N/A'
                campos = [test_desc, test_num, pcb_sn, lower, upper, measure, units, temp_start, temp_end, result]
                if any(c and c.group(1).strip() == "N/A" for c in campos):
                    continue  # Saltar esta prueba

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

    def exportar_excel_formato_personalizado(self, nombre_archivo="tests_excel_formato_personalizado.xlsx"):
        # Agrupar por PCB y número de test
        agrupados = {}
        for test in self.tests:
            clave = (test.serie_pcb, test.test_num, test.descripcion, test.limite_inf, test.limite_sup)
            agrupados.setdefault(clave, []).append(test)

        filas = []
        for clave, tests in agrupados.items():
            pcb, test_num, descripcion, limite_inf, limite_sup = clave
            mediciones = [t.medida for t in tests[:3]]  # Tomar hasta 3 mediciones

            # Mostrar mensajes según cantidad de mediciones
            if len(mediciones) < 3:
                print(f"⚠️ Solo se encontró {len(mediciones)} medición(es) para PCB '{pcb}' Test {test_num}. No se encontraron suficientes archivos para completar Trial2 y Trial3.")
            elif len(mediciones) == 3:
                print(f"✅ Se encontraron las 3 mediciones para PCB '{pcb}' Test {test_num}. Todos los trials están completos.")

            fila = {
                "Test Number": f"Test {test_num}",
                "Test": descripcion.replace(f"Test {test_num} - ", ""),  # Eliminar duplicado si existe
                "LowLimit": limite_inf,
                "HighLimit": limite_sup,
                "PCB Serial Number": pcb,
                "Trial1": mediciones[0] if len(mediciones) > 0 else "",
                "Trial2": mediciones[1] if len(mediciones) > 1 else "",
                "Trial3": mediciones[2] if len(mediciones) > 2 else "",
            }
            filas.append(fila)

        df = pd.DataFrame(filas)
        df.to_excel(nombre_archivo, index=False)
        print(f"✅ Archivo Excel exportado como '{nombre_archivo}' con columnas Trial1, Trial2, Trial3 y Test Number.")

# Función principal
def main():
    ruta = r"C:\Users\3002975\Documents\Py Proyect\leer_log_py"  # Ajustar si es necesario
    procesador = LogProcessor(ruta)
    procesador.procesar_logs()
    procesador.exportar_excel_formato_personalizado()

if __name__ == "__main__":
    main()

