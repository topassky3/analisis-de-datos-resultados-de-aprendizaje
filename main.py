import pandas as pd
import os

# Ruta al archivo de datos de Excel
ruta_archivo = 'Reporte_Juicios_Evaluativos.xls'  # Asegúrate de que este archivo esté en la misma carpeta que el script

# Nombre de la hoja que contiene los datos
nombre_hoja = 'Hoja'  # Cambia esto si tu hoja tiene otro nombre

# Número de filas a saltar antes de los encabezados
# Ajusta este valor según donde estén los encabezados en tu archivo Excel
skip_filas = 12  # Por ejemplo, si los encabezados están en la fila 13

# Lista de números de documento de los estudiantes a verificar
numeros_documento = [
    '1015072940',
    '1025892057',
    '1041632249',
    '1042766722',  
    '1044120917',
    '1044120947',
    '1044121012',
    '1044503792',
    '1044987540'
]

# Directorio de salida para los archivos de resultados
directorio_salida = 'Resultados_Evaluaciones'

# Crear el directorio de salida si no existe
if not os.path.exists(directorio_salida):
    os.makedirs(directorio_salida)

# Cargar los datos desde el archivo Excel
try:
    df = pd.read_excel(ruta_archivo, sheet_name=nombre_hoja, engine='xlrd', skiprows=skip_filas)
except FileNotFoundError:
    print(f"Error: El archivo '{ruta_archivo}' no se encontró en el directorio actual.")
    exit(1)
except ValueError as ve:
    print(f"Error: {ve}")
    exit(1)
except Exception as e:
    print(f"Error al leer el archivo de Excel: {e}")
    exit(1)

# Verificar los nombres de las columnas (opcional)
print("Columnas encontradas en el archivo Excel:")
print(df.columns)

# Normalizar los datos para evitar problemas con espacios y mayúsculas
df['Número de Documento'] = df['Número de Documento'].astype(str).str.strip()
df['Nombre'] = df['Nombre'].astype(str).str.upper().str.strip()
df['Apellidos'] = df['Apellidos'].astype(str).str.upper().str.strip()
df['Juicio de Evaluación'] = df['Juicio de Evaluación'].astype(str).str.upper().str.strip()
df['Resultado de Aprendizaje'] = df['Resultado de Aprendizaje'].astype(str).str.upper().str.strip()
df['Competencia'] = df['Competencia'].astype(str).str.upper().str.strip()

# Filtrar los registros de los estudiantes de interés usando solo el número de documento
df_filtrado = df[df['Número de Documento'].isin(numeros_documento)]

# Verificar si hay registros
if df_filtrado.empty:
    print("No se encontraron registros para los estudiantes especificados.")
    exit(0)

# Definir qué estados indican que fueron evaluados
estados_evaluados = ['APROBADO']

# Agrupar por estudiante y competencia
grupo = df_filtrado.groupby(['Número de Documento', 'Nombre', 'Apellidos', 'Competencia'])

# Crear un diccionario para almacenar los resultados
resultado = {}

for name, group_df in grupo:
    numero_doc, nombre, apellidos, competencia = name
    total_resultados = group_df['Resultado de Aprendizaje'].nunique()
    evaluados = group_df[group_df['Juicio de Evaluación'].str.contains('APROBADO', na=False)]['Resultado de Aprendizaje'].nunique()
    por_evaluar = total_resultados - evaluados

    # Almacenar en el diccionario
    key = numero_doc  # Usamos el número de documento como clave
    if key not in resultado:
        # Guardamos también el nombre y apellidos para usarlo en el archivo de salida
        resultado[key] = {
            'Nombre Completo': f"{nombre.title()} {apellidos.title()}",
            'Competencias': []
        }
    resultado[key]['Competencias'].append({
        'Competencia': competencia,
        'Total Resultados de Aprendizaje': total_resultados,
        'Evaluados': evaluados,
        'Por Evaluar': por_evaluar
    })

# Generar un archivo Excel por estudiante
for numero_doc, datos_estudiante in resultado.items():
    nombre_completo = datos_estudiante['Nombre Completo']
    competencias = datos_estudiante['Competencias']

    # Crear un DataFrame para el estudiante
    data = {
        'Competencia': [],
        'Total Resultados de Aprendizaje': [],
        'Evaluados': [],
        'Por Evaluar': [],
        'Estado': []
    }

    for comp in competencias:
        data['Competencia'].append(comp['Competencia'])
        data['Total Resultados de Aprendizaje'].append(comp['Total Resultados de Aprendizaje'])
        data['Evaluados'].append(comp['Evaluados'])
        data['Por Evaluar'].append(comp['Por Evaluar'])
        if comp['Por Evaluar'] == 0:
            estado = "TODOS los resultados de aprendizaje han sido evaluados."
        else:
            estado = "HAY resultados de aprendizaje POR EVALUAR."
        data['Estado'].append(estado)

    df_estudiante = pd.DataFrame(data)

    # Nombre del archivo Excel de salida
    nombre_archivo_salida = f"{directorio_salida}/{numero_doc}_{nombre_completo.replace(' ', '_')}.xlsx"

    # Escribir el DataFrame en un archivo Excel
    with pd.ExcelWriter(nombre_archivo_salida, engine='openpyxl') as writer:
        df_estudiante.to_excel(writer, sheet_name='Resultados', index=False)

    print(f"Se ha creado el archivo para el estudiante: {nombre_completo} (Documento: {numero_doc})")

print("\nLos resultados se han exportado en la carpeta 'Resultados_Evaluaciones'.")
