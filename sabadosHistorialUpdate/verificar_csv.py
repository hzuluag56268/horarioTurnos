import pandas as pd
import os
from datetime import datetime

def verificar_csv():
    """Verifica el estado actual del archivo CSV."""
    try:
        # Ruta del archivo CSV
        archivo_csv = r"C:\Users\Usuario1\Desktop\cursor\excel_extract\excel_extraction_forschedule\historial_sabados.csv"
        
        print("=== VERIFICACIÓN DEL ARCHIVO CSV ===")
        print("=" * 50)
        
        # Verificar si el archivo existe
        if os.path.exists(archivo_csv):
            print(f"✅ Archivo encontrado: {archivo_csv}")
            
            # Obtener información del archivo
            stat = os.stat(archivo_csv)
            fecha_modificacion = datetime.fromtimestamp(stat.st_mtime)
            tamaño = stat.st_size
            
            print(f"📅 Última modificación: {fecha_modificacion}")
            print(f"📏 Tamaño del archivo: {tamaño} bytes")
            
            # Leer y mostrar el contenido
            df = pd.read_csv(archivo_csv)
            print(f"\n📊 Contenido del archivo:")
            print(f"   • Total de registros: {len(df)}")
            print(f"   • Columnas: {list(df.columns)}")
            
            print(f"\n📋 Datos actuales:")
            print(df.to_string(index=False))
            
            # Verificar valores únicos en la columna de semana
            if 'ultima_semana_trop_sabado' in df.columns:
                valores_unicos = df['ultima_semana_trop_sabado'].dropna().unique()
                print(f"\n📈 Semanas únicas encontradas: {sorted(valores_unicos)}")
                
                # Contar empleados por semana
                for semana in sorted(valores_unicos):
                    count = len(df[df['ultima_semana_trop_sabado'] == semana])
                    print(f"   • Semana {semana}: {count} empleados")
            
        else:
            print(f"❌ Archivo no encontrado: {archivo_csv}")
            print("💡 El archivo no existe o la ruta es incorrecta")
            
    except Exception as e:
        print(f"❌ Error al verificar el archivo: {e}")

if __name__ == "__main__":
    verificar_csv() 