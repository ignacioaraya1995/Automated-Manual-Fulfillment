import pandas as pd
from prettytable import PrettyTable


# Cargar el archivo Excel
# Se asume que el archivo está en el mismo directorio que este script o se proporcionará la ruta completa.
archivo_excel = "roque.xlsx"
df = pd.read_excel(archivo_excel)
# Auditoría
# 1. Contar propiedades (filas) y comparar con un valor dado
num_propiedades = len(df)
# El valor a comparar se solicitará al usuario mediante input en el entorno real de ejecución
valor_comparar = int(input("Ingrese el número de propiedades esperado: "))
comparacion = num_propiedades == valor_comparar

# 2. Verificar si hay celdas vacías en "OWNER FULL NAME"
owner_full_name_vacias = df['OWNER FULL NAME'].isnull().any()

# 3. Verificar si hay celdas vacías en "ADDRESS"
address_vacias = df['ADDRESS'].isnull().any()

# 4. Verificar celdas con valor 0 en "SCORE" o "BUYBOX SCORE"
score_cero = df['SCORE'].eq(0).any()
buybox_score_cero = df['BUYBOX SCORE'].eq(0).any()

# 5. Buscar tags prohibidos en "TAGS"
tags_prohibidos = ['Liti', 'DNC', 'donotmail', 'Takeoff', 'Undeli', 'Return', 'Dead', 'Do Not Mail']
df['TAGS'] = df['TAGS'].fillna('')  # Asegurarse de que no hay NaNs para evitar errores en la búsqueda
tags_encontrados = any(df['TAGS'].str.contains('|'.join(tags_prohibidos), case=False, na=False))

# 6. Buscar valores específicos en "OWNER FULL NAME"
valores_especificos = ['Given Not', 'Record', 'Available', 'Bank', 'Church', 'School', 'Cemetery', 'Not given', 'University', 'College', 'Owner', 'Hospital', 'County', 'City of']
valores_en_owner_full_name = any(df['OWNER FULL NAME'].str.contains('|'.join(valores_especificos), case=False, na=False))

# 7. Comparar "ADDRESS" con "MAILING ADDRESS" cuando "ABSENTEE" sea 1 o 2
absentee_condiciones = df[(df['ADDRESS'] == df['MAILING ADDRESS']) & df['ABSENTEE'].isin([1, 2])]

# 8. Revisar "ACTION PLANS" para casos urgentes con "SCORE" menor a 585
casos_urgentes = df[(df['ACTION PLANS'].str.contains('urgent', case=False, na=False)) & (df['SCORE'] < 585)]

# Análisis
# 1. Tabla de "ACTION PLANS"
tabla_action_plans = df['ACTION PLANS'].value_counts().reset_index()
tabla_action_plans.columns = ['Action Plan', 'Frecuencia']
tabla_action_plans['Porcentaje'] = (tabla_action_plans['Frecuencia'] / num_propiedades) * 100

# 2. Tabla de "ESTIMATED MARKET VALUE"
bins = [0, 100000, 200000, 300000, 400000, 500000, 600000, 700000, 800000, 900000, 1000000, float('inf')]
labels = ['0-100k', '100k-200k', '200k-300k', '300k-400k', '400k-500k', '500k-600k', '600k-700k', '700k-800k', '800k-900k', '900k-1M', '1M+']
df['Market Value Range'] = pd.cut(df['ESTIMATED MARKET VALUE'], bins=bins, labels=labels, right=False)
tabla_market_value = df['Market Value Range'].value_counts().sort_index().reset_index()
tabla_market_value.columns = ['Market Value Range', 'Número de Propiedades']
tabla_market_value['Porcentaje'] = (tabla_market_value['Número de Propiedades'] / num_propiedades) * 100

# Resultados (Estos se adaptarían a print o retorno de valores según la necesidad)
# Nota: La ejecución directa y la interacción con el usuario mediante input están comentadas dado que no se pueden ejecutar en este entorno.
print(f"Comparación de cantidad de propiedades con valor ingresado: {comparacion}")
print(f"Celdas vacías en 'OWNER FULL NAME': {'Sí' if owner_full_name_vacias else 'No'}")
print(f"Celdas vacías en 'ADDRESS': {'Sí' if address_vacias else 'No'}")
print(f"Celdas con valor 0 en 'SCORE' o 'BUYBOX SCORE': {'Sí' if score_cero or buybox_score_cero else 'No'}")
print(f"Tags prohibidos encontrados en 'TAGS': {'Sí' if tags_encontrados else 'No'}")
print(f"Valores específicos en 'OWNER FULL NAME': {'Sí' if valores_en_owner_full_name else 'No'}")
print(f"Propiedades con 'ADDRESS' igual a 'MAILING ADDRESS' y 'ABSENTEE' 1 o 2: {'Sí' if not absentee_condiciones.empty else 'No'}")
print(f"Casos urgentes con 'SCORE' menor a 585: {'Sí' if not casos_urgentes.empty else 'No'}")
print(tabla_action_plans)
print(tabla_market_value)

import pandas as pd
from prettytable import PrettyTable

def formatear_numero(n):
    """Formatea el número con comas como separadores de miles."""
    return "{:,.0f}".format(n).replace(",", ".")

def formatear_porcentaje(p):
    """Formatea el porcentaje con dos decimales."""
    return "{:.2f}%".format(p * 100)

def calcular_acumulado(serie):
    """Calcula el porcentaje acumulado de una serie."""
    return serie.cumsum()

def analizar_y_presentar(df):
    total_props = len(df)
    
    # Análisis y presentación para ZIP
    print("Análisis de ZIP:")
    analizar_columna(df, 'ZIP', total_props)
    
    # Análisis y presentación para SCORE
    print("\nAnálisis de SCORE:")
    score_bins = range(0, 1001, 50)  # Ajusta según sea necesario
    score_labels = [f"{i}-{i+49}" for i in score_bins[:-1]]
    df['SCORE Range'] = pd.cut(df['SCORE'], bins=score_bins, labels=score_labels, right=False)
    analizar_columna(df, 'SCORE Range', total_props, 'Rango de SCORE')
    
    # Análisis y presentación para PROPERTY TYPE
    print("\nAnálisis de PROPERTY TYPE:")
    analizar_columna(df, 'PROPERTY TYPE', total_props)

def analizar_columna(df, columna, total_props, nombre_columna=None):
    if not nombre_columna:
        nombre_columna = columna
    tabla = PrettyTable()
    tabla.field_names = [nombre_columna, "Número de Propiedades", "Porcentaje", "Acumulado"]
    counts = df[columna].value_counts().sort_index()
    porcentajes = counts / total_props
    acumulado = calcular_acumulado(porcentajes)
    
    for valor, count in counts.items():
        porcentaje = porcentajes[valor]
        acum = acumulado[valor]
        tabla.add_row([valor, formatear_numero(count), formatear_porcentaje(porcentaje), formatear_porcentaje(acum)])
    
    print(tabla)

# Supongamos que df es tu DataFrame cargado de 'roque.xlsx'
analizar_y_presentar(df)
