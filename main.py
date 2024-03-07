import pandas as pd
import os
from tqdm import tqdm

# Función para filtrar filas basadas en tags prohibidos
def filter_prohibited_tags(df, column_name, prohibited_tags):
    pattern = '|'.join(prohibited_tags)  # Crear un patrón de expresión regular
    return df[~df[column_name].str.contains(pattern, case=False, na=False)]

# Función para filtrar filas basadas en nombres no deseados
def filter_unwanted_names(df, column_name, unwanted_names):
    pattern = '|'.join(unwanted_names)
    return df[~df[column_name].str.contains(pattern, case=False, na=False)]

# Función principal para procesar los archivos
def process_excel_files(folder_path):
    deletion_messages = []
    files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]

    for file in tqdm(files, desc="Procesando archivos"):
        file_path = os.path.join(folder_path, file)
        df = pd.read_excel(file_path)

        # Eliminar la primera columna
        df.drop(df.columns[0], axis=1, inplace=True)
        
        # Igualar celdas vacías en "OWNER LAST NAME"
        df['OWNER LAST NAME'].fillna(df['OWNER FULL NAME'], inplace=True)
        
        # Ordenar valores en orden decreciente
        df.sort_values(by=['BUYBOX SCORE', 'LIKELY DEAL SCORE', 'SCORE'], ascending=False, inplace=True)

        # Eliminar filas con tags prohibidos
        prohibited_tags = ["Liti", "DNC", "donotmail", "Takeoff", "Undeli", "Return", "Dead", "Do Not Mail"]
        df = filter_prohibited_tags(df, 'TAGS', prohibited_tags)
        
        # Eliminar filas donde "ACTION PLANS" están en blanco
        df = df[df['ACTION PLANS'].notna()]

        # Igualar columnas vacías
        df['MAILING ADDRESS'].fillna(df['ADDRESS'], inplace=True)
        df['MAILING ZIP'].fillna(df['ZIP'], inplace=True)
        
        # Remover duplicados basados en combinaciones de columnas
        df.drop_duplicates(subset=['MAILING ADDRESS', 'MAILING ZIP'], inplace=True)
        df.drop_duplicates(subset=['OWNER FULL NAME', 'ADDRESS', 'ZIP'], inplace=True)

        # Remover unwanted "OWNER FULL NAMES"
        unwanted_names = ["Given Not", "Record", "Available", "Bank", "Church", "School", "Cemetery", "Not given", "University", "College", "Owner", "Hospital", "County", "City of"]
        df = filter_unwanted_names(df, 'OWNER FULL NAME', unwanted_names)

        # Guardar el archivo modificado
        output_file_path = f"output - {file}"
        df.to_excel(output_file_path, index=False)

        deletion_messages.append(f"Archivo {file} procesado y guardado.")

    for message in deletion_messages:
        print(message)

# Ejemplo de llamada a la función:
process_excel_files("data")
