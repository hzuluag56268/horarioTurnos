import pandas as pd

# Leer el archivo generado
df = pd.read_excel('horario_descansos_julio.xlsx')

print("=== VERIFICACIÓN DEL HORARIO GENERADO ===")
print(f"Dimensiones: {df.shape}")
print(f"Columnas: {df.columns.tolist()}")

print("\nPrimeras 3 filas:")
print(df.head(3))

print("\n=== ANÁLISIS DE DESCANSO POR EMPLEADO ===")
for idx, row in df.iterrows():
    empleado = row['SIGLA ATCO']
    desc_count = sum(1 for col in df.columns if col.startswith(('MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT', 'SUN')) 
                    and row[col] == 'DESC')
    trop_count = sum(1 for col in df.columns if col.startswith(('MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT', 'SUN')) 
                    and row[col] == 'TROP')
    
    print(f"Empleado {row['No.']} ({empleado}): DESC={desc_count}, TROP={trop_count}")

print("\n=== VERIFICACIÓN DE DOMINGOS ===")
# Verificar que ningún empleado tenga descanso en domingo
domingos = [col for col in df.columns if col.startswith('SUN')]
print(f"Domingos en el mes: {domingos}")

for domingo in domingos:
    descansos_domingo = df[df[domingo].isin(['DESC', 'TROP'])]
    if len(descansos_domingo) > 0:
        print(f"⚠️  PROBLEMA: {len(descansos_domingo)} empleados tienen descanso en {domingo}")
        for idx, row in descansos_domingo.iterrows():
            print(f"   - Empleado {row['No.']} ({row['SIGLA ATCO']}): {row[domingo]}")
    else:
        print(f"✅ Correcto: Ningún empleado descansa en {domingo}")

print("\n=== EJEMPLO DE DISTRIBUCIÓN SEMANAL ===")
# Mostrar la primera semana de un empleado
empleado_ejemplo = df.iloc[0]
print(f"Empleado {empleado_ejemplo['No.']} ({empleado_ejemplo['SIGLA ATCO']}):")
for col in df.columns[2:9]:  # Primera semana
    valor = empleado_ejemplo[col]
    print(f"  {col}: {valor if pd.notna(valor) else 'TRABAJO'}") 