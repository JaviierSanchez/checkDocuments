# Check Documents

Este script de Python verifica la existencia de documentos listados en un archivo Excel dentro de una carpeta específica. Genera reportes separados para documentos encontrados y no encontrados, y organiza los resultados en una carpeta con el mismo nombre que el archivo Excel original.

## Características

* Lee rutas de documentos desde un archivo Excel (`.xlsx`).
* Verifica la existencia de los documentos en una carpeta y sus subcarpetas.
* Genera dos archivos Excel de salida: `encontrados.xlsx` y `no_encontrados.xlsx`.
* Crea una carpeta de resultados con el nombre del archivo Excel de entrada.
* Muestra un resumen de la búsqueda en la consola.
* Manejo de rutas relativas y absolutas.
* Normalización de rutas para evitar problemas de mayúsculas/minúsculas y barras diagonales.
* Adaptabilidad a diferentes carpetas de documentos.

## Requisitos

* Python 3.x
* Librería `pandas` (`pip install pandas`)

## Uso

1.  Clona el repositorio.
2.  Asegúrate de tener un archivo Excel (`.xlsx`) con una columna llamada "RUTA" que contenga las rutas de los documentos a verificar.
3.  Modifica las rutas `excel_path` y `folder_path` en el script `check_Documents.py` para que apunten a tu archivo Excel y la carpeta de documentos, respectivamente.
4.  Ejecuta el script: `python check_Documents.py`
5.  Los resultados se generarán en una carpeta con el mismo nombre que el archivo Excel, dentro de la carpeta "Resultado" del script.

## Ejemplo de uso

Suponiendo que tienes un archivo Excel llamado "Documentos.xlsx" con rutas de documentos en la columna "RUTA" y los documentos se encuentran en la carpeta "Documentos_A_Verificar", debes modificar las siguientes líneas al principio del script:

```python
excel_path = "C:/ruta/a/Documentos.xlsx"
folder_path = "C:/ruta/a/Documentos_A_Verificar"
