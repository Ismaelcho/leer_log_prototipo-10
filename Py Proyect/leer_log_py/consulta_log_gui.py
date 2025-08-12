import tkinter as tk
from tkinter import ttk

# Simulación de datos de LogTest
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
# Datos de ejemplo
tests = [
    LogTest("log1.txt", 101, "Voltage Test", "PCB A", "1.0", "5.0", "3.2", "V", "25", "30", "Pass"),
    LogTest("log2.txt", 102, "Current Test", "PCB B", "0.5", "2.0", "1.8", "A", "20", "25", "Fail"),
    LogTest("log3.txt", 103, "Resistance Test", "PCB C", "10", "100", "55", "Ohm", "22", "28", "Pass"),
]

def filtrar_tests():
    descripcion = entry_desc.get().lower()
    test_num = entry_num.get()
    pcb = entry_pcb.get().lower()
    resultado = entry_result.get().lower()

    filtrados = []
    for t in tests:
        if descripcion and descripcion not in t.descripcion.lower():
            continue
        if test_num and str(t.test_num) != test_num:
            continue
        if pcb and pcb not in t.serie_pcb.lower():
            continue
        if resultado and resultado not in t.resultado.lower():
            continue
        filtrados.append(t)

    mostrar_resultados(filtrados)

def mostrar_resultados(lista):
    for row in tree.get_children():
        tree.delete(row)
    for t in lista:
        tree.insert("", "end", values=(t.test_num, t.descripcion, t.serie_pcb, t.limite_inf,
                                       t.limite_sup, t.medida, t.unidades, t.temp_ini,
                                       t.temp_fin, t.resultado))

# Crear ventana principal
root = tk.Tk()
root.title("Consulta de Tests de Log")

# Campos de entrada
tk.Label(root, text="Descripción:").grid(row=0, column=0)
entry_desc = tk.Entry(root)
entry_desc.grid(row=0, column=1)

tk.Label(root, text="Número de Test:").grid(row=1, column=0)
entry_num = tk.Entry(root)
entry_num.grid(row=1, column=1)

tk.Label(root, text="Serial PCB:").grid(row=2, column=0)
entry_pcb = tk.Entry(root)
entry_pcb.grid(row=2, column=1)

tk.Label(root, text="Resultado:").grid(row=3, column=0)
entry_result = tk.Entry(root)
entry_result.grid(row=3, column=1)

tk.Button(root, text="Buscar", command=filtrar_tests).grid(row=4, column=0, columnspan=2)

# Tabla de resultados
columns = ("Test Num", "Descripción", "PCB", "Límite Inf", "Límite Sup", "Medida", "Unidades", "Temp Ini", "Temp Fin", "Resultado")
tree = ttk.Treeview(root, columns=columns, show="headings")
for col in columns:
    tree.heading(col, text=col)
tree.grid(row=5, column=0, columnspan=2)

root.mainloop()
