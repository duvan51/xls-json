import openpyxl
import json
from tkinter import Tk, filedialog

# Ventana de selección de archivo
root = Tk()
root.withdraw()  # Oculta la ventana principal

# Solicitar al usuario que seleccione un archivo Excel
excel_file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])

# Cargar el archivo Excel
workbook = openpyxl.load_workbook(excel_file_path)

# Obtener la primera hoja (sheet)
sheet = workbook.active

# Obtener los encabezados de las columnas
headers = [cell.value for cell in sheet[1]]

# Buscar los índices de las columnas "codigo" y "costo"
codigo_index = headers.index("procod")
costo_index = headers.index("costo")
nombre_index = headers.index("pronom")
existencias_index = headers.index("existen")

# Convertir las filas de datos a formato JSON
data = []
for row in sheet.iter_rows(min_row=2, values_only=True):
    codigo = row[codigo_index]
    costo = row[costo_index]
    nombre = row[nombre_index]
    existencias = row[existencias_index]
    data.append({"codigo": codigo, "costo": costo, "nombre": nombre, "existencias": existencias})

# Guardar en formato JSON
json_file_path = 'ruta_a_guardar_tu_archivo.json'
with open(json_file_path, 'w') as json_file:
    json.dump(data, json_file, indent=2)

print('Archivo JSON creado exitosamente:', json_file_path)
