import pandas as pd

# Leer el archivo Excel
df = pd.read_excel('FormatodeSalidaRequerido.xlsx')

print("=== ESTRUCTURA DEL ARCHIVO DE FORMATO REQUERIDO ===")
print(f"Dimensiones: {df.shape}")
print(f"Columnas: {df.columns.tolist()}")
print("\nPrimeras 5 filas:")
print(df.head())

print("\nTipos de datos:")
print(df.dtypes)

print("\nValores únicos en la primera columna (empleados):")
print(df.iloc[:, 0].unique())

print("\nEjemplo de valores en las columnas de días:")
for col in df.columns[2:7]:  # Primeras 5 columnas de días
    print(f"{col}: {df[col].unique()}") 