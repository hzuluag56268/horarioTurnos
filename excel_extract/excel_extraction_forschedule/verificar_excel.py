import pandas as pd
import os

def verificar_archivo_excel(nombre_archivo):
    """Verifica el contenido del archivo Excel"""
    print(f"🔍 VERIFICANDO ARCHIVO EXCEL: {nombre_archivo}")
    
    # Verificar si el archivo existe
    if not os.path.exists(nombre_archivo):
        print(f"❌ El archivo {nombre_archivo} no existe")
        return
    
    try:
        # Leer el archivo Excel
        df = pd.read_excel(nombre_archivo)
        
        print(f"✅ Archivo leído exitosamente")
        print(f"📊 Dimensiones: {df.shape[0]} filas x {df.shape[1]} columnas")
        
        # Mostrar información de columnas
        print(f"\n📋 Columnas:")
        for i, col in enumerate(df.columns):
            print(f"  {i+1}. {col}")
        
        # Mostrar primeras filas
        print(f"\n📋 Primeras filas:")
        print(df.head(10).to_string(index=False))
        
        # Análisis de datos
        print(f"\n📊 Análisis de datos:")
        
        # Contar valores por columna
        for col in df.columns[2:]:  # Excluir No. y SIGLA ATCO
            valores = df[col].value_counts()
            print(f"\n  {col}:")
            for valor, count in valores.items():
                print(f"    {valor}: {count}")
        
        # Verificar restricciones
        print(f"\n✅ Verificación de restricciones:")
        
        # Verificar domingos
        domingos = [col for col in df.columns if col.startswith('SUN')]
        for domingo in domingos:
            personas_trabajando = sum(1 for valor in df[domingo] if valor is None)
            print(f"  {domingo}: {personas_trabajando}/10 trabajando")
        
        # Verificar días laborables
        dias_laborables = [col for col in df.columns if not col.startswith('SUN') and col not in ['No.', 'SIGLA ATCO']]
        for dia in dias_laborables:
            personas_trabajando = sum(1 for valor in df[dia] if valor is None)
            print(f"  {dia}: {personas_trabajando} trabajando, {10-personas_trabajando} descansando")
        
        # Verificar descansos por empleado
        print(f"\n👥 Descansos por empleado:")
        for idx, empleado in enumerate(df['SIGLA ATCO']):
            descansos = sum(1 for col in df.columns[2:] if df.iloc[idx][col] in ['DESC', 'TROP'])
            sabados = sum(1 for col in df.columns[2:] if col.startswith('SAT') and df.iloc[idx][col] in ['DESC', 'TROP'])
            print(f"  {empleado}: {descansos} descansos totales, {sabados} sábados")
        
        print(f"\n🎉 Verificación completada exitosamente!")
        
    except Exception as e:
        print(f"❌ Error al leer el archivo: {e}")

def main():
    archivos_excel = [
        'horario_primera_semana_julio.xlsx',
        'horario_heuristico_semanal_final_julio.xlsx',
        'horario_heuristico_semanal_mejorado_julio.xlsx',
        'horario_heuristico_semanal_julio.xlsx'
    ]
    
    print("=== VERIFICACIÓN DE ARCHIVOS EXCEL ===")
    
    for archivo in archivos_excel:
        if os.path.exists(archivo):
            print(f"\n{'='*50}")
            verificar_archivo_excel(archivo)
        else:
            print(f"\n❌ Archivo {archivo} no encontrado")

if __name__ == "__main__":
    main() 