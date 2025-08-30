import pandas as pd
import numpy as np

# Leer el archivo Excel
df = pd.read_excel('FormatodeSalidaRequerido.xlsx')

print("=== ANÁLISIS DETALLADO DEL FORMATO REQUERIDO ===")

# Filtrar solo las filas de empleados (excluir la última fila que parece ser un título)
empleados_df = df[df['No.'].notna() & (df['No.'] != 'Coordinador Grupo Regional Servicios de Tránsito Aéreo')].copy()

print(f"Empleados válidos: {len(empleados_df)}")

# Obtener columnas de días (excluir 'No.', 'SIGLA ATCO', 'SIGLAATCO' y columnas 'Unnamed')
columnas_dias = [col for col in df.columns if '-' in col and 'Unnamed' not in col]
print(f"\nColumnas de días encontradas: {columnas_dias}")

# Analizar códigos únicos en las columnas de días
todos_codigos = []
for col in columnas_dias:
    valores = empleados_df[col].dropna().unique()
    todos_codigos.extend(valores)

codigos_unicos = list(set(todos_codigos))
print(f"\nCódigos únicos encontrados: {sorted(codigos_unicos)}")

# Contar frecuencia de cada código
print("\nFrecuencia de códigos:")
for codigo in sorted(codigos_unicos):
    count = sum(empleados_df[col].value_counts().get(codigo, 0) for col in columnas_dias)
    print(f"  {codigo}: {count} veces")

# Mostrar ejemplo de un empleado completo
print(f"\nEjemplo - Empleado {empleados_df.iloc[0]['No.']} ({empleados_df.iloc[0]['SIGLA ATCO']}):")
for col in columnas_dias:
    valor = empleados_df.iloc[0][col]
    if pd.notna(valor):
        print(f"  {col}: {valor}")

# Verificar si DESC y TROP aparecen exactamente 2 veces por semana
print("\n=== VERIFICACIÓN DE DESC Y TROP ===")
for idx, empleado in empleados_df.iterrows():
    desc_count = sum(1 for col in columnas_dias if empleado[col] == 'DESC')
    trop_count = sum(1 for col in columnas_dias if empleado[col] == 'TROP')
    print(f"Empleado {empleado['No.']} ({empleado['SIGLA ATCO']}): DESC={desc_count}, TROP={trop_count}") 