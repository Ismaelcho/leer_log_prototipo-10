import pandas as pd
from pathlib import Path
import re

#TestResult → Representa cada test individual extraído del log.
class TestResult:
    def __init__(self, archivo, test_num, descripcion, serie_pcb, limite_inferior, limite_superior, medida, unidades, temp_inicio, temp_fin, resultado):
        self.archivo = archivo
        self.test_num = test_num
        self.descripcion = descripcion
        self.serie_pcb = serie_pcb
        self.limite_inferior = limite_inferior
        self.limite_superior = limite_superior
        self.medida = medida
        self.unidades = unidades
        self.temp_inicio = temp_inicio
        self.temp_fin = temp_fin
        self.resultado = resultado

    def to_dict(self):
        return {
            "Archivo": self.archivo,
            "Test Nº": self.test_num,
            "Descripción": self.descripcion,
            "Serie PCB": self.serie_pcb,
            "Límite Inferior": self.limite_inferior,
            "Límite Superior": self.limite_superior,
            "Medida": self.medida,
            "Unidades": self.unidades,
            "Temp Inicio": self.temp_inicio,
            "Temp Fin": self.temp_fin,
            "Resultado": self.resultado
        }

#LogFile → Representa un archivo de log y contiene métodos para analizar su contenido.
class LogFile:
    def __init__(self, path):
        self.path = path
        self.tests = []

    def parse(self):
        with self.path.open("r", encoding="utf-8") as f:
            content = f.read()

        bloques = re.split(r"(?=Test Description:)", content)
        for bloque in bloques:
            if "Test Description:" in bloque:
                self.tests.append(self._parse_test_block(bloque))

    def _parse_test_block(self, bloque):
        def get(pattern):  # helper interno
            match = re.search(pattern, bloque)
            return match.group(1).strip() if match else None

        return TestResult(
            archivo=self.path.name,
            test_num=get(r"Test (\d+)"),
            descripcion=get(r"Test Description:\s*(.*)"),
            serie_pcb=get(r"PCB Serial Number:\s*(.*)"),
            limite_inferior=get(r"Test Lower Limit:\s*(.*)"),
            limite_superior=get(r"Test Upper Limit:\s*(.*)"),
            medida=get(r"Test Measurement:\s*(.*)"),
            unidades=get(r"Units:\s*(.*)"),
            temp_inicio=get(r"Starting Temperature.*:\s*(.*)"),
            temp_fin=get(r"Ending Temperature.*:\s*(.*)"),
            resultado=get(r"Test Result:\s*(.*)")
        )

#LogExtractor → Clase principal que gestiona todo el proceso.
class LogExtractor:
    def __init__(self, carpeta_logs):
        self.carpeta_logs = Path(carpeta_logs)
        self.test_results = []

    def run(self):
        if not self.carpeta_logs.exists():
            print("Carpeta no encontrada:", self.carpeta_logs)
            return

        for log_path in self.carpeta_logs.glob("*.txt"):
            log_file = LogFile(log_path)
            log_file.parse()
            self.test_results.extend(log_file.tests)

    def export_to_excel(self, output_file="tests_extraidos.xlsx"):
        df = pd.DataFrame([test.to_dict() for test in self.test_results])
        df.to_excel(output_file, index=False)
        print(f"Todos los tests han sido extraídos y guardados en '{output_file}'.")


# --- Ejecución principal ---
if __name__ == "__main__":
    extractor = LogExtractor(r"C:\Users\3002975\Documents\Py Proyect\leer_log_py")
    extractor.run()
    extractor.export_to_excel()
