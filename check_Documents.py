import os
import pandas as pd

# Configuración: Rutas de archivos y carpetas
excel_path = "Ruta_Absoluta_Excel"
folder_path = "Ruta_Absoluta_Carpeta_Documentos"
output_folder_base = "Ruta_Absoluta_Donde_Se_Guardan_Resultados" #Ruta Base de la carpeta de resultados

print("Cargando archivo Excel...")
# Cargar la hoja específica del Excel (índice 9 -> la décima hoja)
sheet_name = "estado tipos documentos"
df = pd.read_excel(excel_path, sheet_name=sheet_name)
print(f"Archivo Excel '{excel_path}' cargado correctamente.")

# Verificar que la columna "RUTA" existe en la hoja
columna_ruta = "RUTA"
if columna_ruta not in df.columns:
    raise ValueError(f"La columna '{columna_ruta}' no se encuentra en la hoja '{sheet_name}'.")

print(f"Procesando {len(df)} documentos...")

# Lista para guardar resultados
resultados = []

# Recorrer la lista de rutas de documentos
for index, ruta_documento in enumerate(df[columna_ruta]):
    print(f"[{index+1}/{len(df)}] Verificando documento: {ruta_documento}") # Primera línea

    if not isinstance(ruta_documento, str):  # Si la celda está vacía o no es texto, saltarla
        resultados.append({"Documento": ruta_documento, "Encontrado": "No (Ruta no válida)"})
        print("  Resultado: No (Ruta no válida)") #Segunda línea
        continue

    encontrado = False
    ruta_encontrada = "No encontrado"

    # Obtener el nombre de la carpeta de documentos de folder_path
    folder_name = os.path.basename(folder_path)

    # Construir la cadena de búsqueda dinámica
    search_string = f"Documentos\\{folder_name}\\"

    # Extraer la parte relevante de la ruta del Excel
    ruta_relativa = ruta_documento.replace(search_string, "", 1).strip()  # Elimina la carpeta dinámicamente

    # Construir la ruta completa del archivo
    ruta_completa = os.path.normcase(os.path.normpath(os.path.join(folder_path, ruta_relativa)))

    # Verificar si el archivo existe
    if os.path.isfile(ruta_completa):
        encontrado = True
        ruta_encontrada = ruta_completa
        print(f"✓  Resultado: Encontrado: {ruta_completa}") #Segunda línea
    else:
        encontrado = False
        print(f"✗  Resultado: NO encontrado: {ruta_documento}") #Segunda línea

    # Guardar el resultado
    resultados.append({"Documento": ruta_documento, "Encontrado": "Sí" if encontrado else "No", "Ruta Exacta": ruta_encontrada})

# Convertir resultados a un DataFrame y exportar a Excel
df_resultados = pd.DataFrame(resultados)

# Crear la carpeta de resultados
excel_filename_without_extension = os.path.splitext(os.path.basename(excel_path))[0]
output_folder = os.path.join(
    output_folder_base,
    excel_filename_without_extension,
)  # Ruta de la carpeta de resultados
os.makedirs(output_folder, exist_ok=True)  # Crea la carpeta si no existe

# Filtrar los resultados en encontrados y no encontrados
df_encontrados = df_resultados[df_resultados["Encontrado"] == "Sí"]
df_no_encontrados = df_resultados[df_resultados["Encontrado"] == "No"]

# Guardar los DataFrames en archivos Excel separados
encontrados_excel_path = os.path.join(output_folder, "encontrados.xlsx")
no_encontrados_excel_path = os.path.join(output_folder, "no_encontrados.xlsx")

df_encontrados.to_excel(encontrados_excel_path, index=False)
df_no_encontrados.to_excel(no_encontrados_excel_path, index=False)

# Calcular y mostrar el resumen
total_documentos = len(df_resultados)
documentos_encontrados = len(df_encontrados)
documentos_no_encontrados = len(df_no_encontrados)

print("\n--- Resumen de la búsqueda ---")
print(f"Total de documentos procesados: {total_documentos}")
print(f"Documentos encontrados: {documentos_encontrados}")
print(f"Documentos no encontrados: {documentos_no_encontrados}")
print(
    f"\nBúsqueda completada. Los resultados se han guardado en la carpeta '{output_folder}'."
)