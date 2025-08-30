import pandas as pd
import os

def verificar_archivo_real():
    """Verifica el contenido real del archivo CSV."""
    try:
        # Ruta del archivo CSV
        archivo_csv = r"C:\Users\Usuario1\Desktop\cursor\excel_extract\excel_extraction_forschedule\historial_sabados.csv"
        
        print("=== VERIFICACIÓN DEL ARCHIVO REAL ===")
        print("=" * 50)
        
        # Verificar si el archivo existe
        if os.path.exists(archivo_csv):
            print(f"✅ Archivo encontrado: {archivo_csv}")
            
            # Obtener información del archivo
            stat = os.stat(archivo_csv)
            fecha_modificacion = stat.st_mtime
            tamaño = stat.st_size
            
            print(f"📅 Última modificación: {fecha_modificacion}")
            print(f"📏 Tamaño del archivo: {tamaño} bytes")
            
            # Leer el archivo como texto para ver el contenido real
            with open(archivo_csv, 'r', encoding='utf-8') as f:
                contenido = f.read()
            
            print(f"\n📄 CONTENIDO REAL DEL ARCHIVO:")
            print("=" * 50)
            print(contenido)
            print("=" * 50)
            
            # También leer con pandas para comparar
            df = pd.read_csv(archivo_csv)
            print(f"\n📊 CONTENIDO CON PANDAS:")
            print(df.to_string(index=False))
            
            # Verificar valores únicos
            if 'ultima_semana_trop_sabado' in df.columns:
                valores_unicos = df['ultima_semana_trop_sabado'].dropna().unique()
                print(f"\n📈 Valores únicos encontrados: {sorted(valores_unicos)}")
                
                # Contar por valor
                for valor in sorted(valores_unicos):
                    count = len(df[df['ultima_semana_trop_sabado'] == valor])
                    print(f"   • Valor {valor}: {count} empleados")
            
        else:
            print(f"❌ Archivo no encontrado: {archivo_csv}")
            
    except Exception as e:
        print(f"❌ Error al verificar el archivo: {e}")

if __name__ == "__main__":
    verificar_archivo_real() 