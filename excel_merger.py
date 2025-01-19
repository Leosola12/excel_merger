import pandas as pd
import os
from tkinter import Tk
from tkinter.filedialog import askdirectory, asksaveasfilename

# Ocultar la ventana principal de Tkinter
Tk().withdraw()

# Seleccionar carpeta que contiene los archivos Excel
print("Selecciona la carpeta que contiene los archivos Excel:")
folder_path = askdirectory(title="Seleccionar carpeta de entrada")
if not folder_path:
    print("No se seleccionó ninguna carpeta. Saliendo del script.")
    exit()

# Lista para almacenar los datos de cada archivo
dataframes = []

# Iterar sobre todos los archivos Excel en la carpeta
for filename in os.listdir(folder_path):
    if filename.endswith('.xlsx') or filename.endswith('.xls'):
        # Leer cada archivo y agregar a la lista
        file_path = os.path.join(folder_path, filename)
        df = pd.read_excel(file_path)
        dataframes.append(df)

# Combinar todos los DataFrames en uno solo
combined_df = pd.concat(dataframes, ignore_index=True)

# Seleccionar ubicación y nombre del archivo de salida
print("Selecciona el lugar y nombre para guardar el archivo consolidado:")
output_file = asksaveasfilename(
    title="Guardar archivo consolidado",
    defaultextension=".xlsx",
    filetypes=[("Archivos Excel", "*.xlsx")]
)
if not output_file:
    print("No se seleccionó ninguna ubicación para guardar el archivo. Saliendo del script.")
    exit()

# Guardar el DataFrame combinado en el archivo seleccionado
combined_df.to_excel(output_file, index=False)

print(f"Archivos consolidados guardados en: {output_file}")
