import pandas as pd
import os
from datetime import datetime

def corregir_formato_csv():
    """Corrige el formato del archivo CSV para que los números sean enteros."""
    try:
        # Ruta del archivo CSV
        archivo_csv = r"C:\Users\Usuario1\Desktop\cursor\excel_extract\excel_extraction_forschedule\historial_sabados.csv"
        
        print("=== CORRECCIÓN DE FORMATO CSV ===")
        print("=" * 50)
        
        if not os.path.exists(archivo_csv):
            print(f"❌ Archivo no encontrado: {archivo_csv}")
            return
        
        # Leer el archivo actual
        df = pd.read_csv(archivo_csv)
        print("📊 Archivo actual:")
        print(df.to_string(index=False))
        
        # Convertir la columna a enteros
        df['ultima_semana_trop_sabado'] = pd.to_numeric(df['ultima_semana_trop_sabado'], errors='coerce')
        
        # Guardar con formato de enteros
        df.to_csv(archivo_csv, index=False, float_format='%.0f')
        
        print(f"\n✅ Archivo corregido y guardado")
        print(f"📁 Ubicación: {archivo_csv}")
        
        # Verificar el resultado
        df_corregido = pd.read_csv(archivo_csv)
        print(f"\n📋 Resultado final:")
        print(df_corregido.to_string(index=False))
        
        # Mostrar estadísticas
        empleados_con_semana = df_corregido['ultima_semana_trop_sabado'].notna().sum()
        empleados_sin_semana = df_corregido['ultima_semana_trop_sabado'].isna().sum()
        
        print(f"\n📈 Estadísticas:")
        print(f"   • Total de empleados: {len(df_corregido)}")
        print(f"   • Con semana asignada: {empleados_con_semana}")
        print(f"   • Sin semana asignada: {empleados_sin_semana}")
        
        # Mostrar semanas únicas
        valores_unicos = df_corregido['ultima_semana_trop_sabado'].dropna().unique()
        print(f"   • Semanas únicas: {sorted(valores_unicos)}")
        
    except Exception as e:
        print(f"❌ Error al corregir el archivo: {e}")

if __name__ == "__main__":
    corregir_formato_csv() 