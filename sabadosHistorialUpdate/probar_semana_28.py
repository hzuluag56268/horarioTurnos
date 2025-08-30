import pandas as pd
import openpyxl
import re
import os
from datetime import datetime

def extraer_numero_semana(nombre_archivo):
    """Extrae el número de semana del nombre del archivo Excel."""
    match = re.search(r'semana_(\d+)', nombre_archivo)
    if match:
        return int(match.group(1))
    return None

def extraer_iniciales_con_trop(archivo_excel):
    """Extrae las iniciales de empleados que tienen 'TROP' en columnas SAT."""
    try:
        print(f"🔍 Procesando archivo: {archivo_excel}")
        
        # Extraer número de semana
        numero_semana = extraer_numero_semana(archivo_excel)
        print(f"📅 Número de semana detectado: {numero_semana}")
        
        # Leer el archivo Excel
        excel_file = pd.ExcelFile(archivo_excel)
        hojas = excel_file.sheet_names
        print(f"📋 Hojas encontradas: {hojas}")
        
        iniciales_con_trop = []
        
        for hoja in hojas:
            print(f"\n📄 Procesando hoja: {hoja}")
            
            df = pd.read_excel(archivo_excel, sheet_name=hoja)
            print(f"📊 Columnas en la hoja: {list(df.columns)}")
            
            # Buscar columnas que empiecen con 'SAT'
            columnas_sat = [col for col in df.columns if str(col).startswith('SAT')]
            print(f"🔍 Columnas SAT encontradas: {columnas_sat}")
            
            if columnas_sat:
                for col_sat in columnas_sat:
                    print(f"\n🔍 Buscando 'TROP' en columna: {col_sat}")
                    
                    # Buscar filas con 'TROP'
                    trop_filas = df[df[col_sat].astype(str).str.contains('TROP', case=False, na=False)]
                    
                    if not trop_filas.empty:
                        print(f"✅ Encontradas {len(trop_filas)} filas con TROP en {col_sat}")
                        
                        for idx, fila in trop_filas.iterrows():
                            inicial = None
                            
                            # Buscar iniciales en las primeras columnas
                            for col in df.columns[:5]:
                                valor = str(fila[col]).strip()
                                if len(valor) <= 4 and valor.isalpha():
                                    inicial = valor
                                    break
                            
                            if inicial:
                                iniciales_con_trop.append({
                                    'inicial': inicial,
                                    'columna_sat': col_sat,
                                    'fila': idx + 1,
                                    'valor_completo': str(fila[col_sat])
                                })
                                print(f"  ✅ Encontrado TROP: {inicial} en {col_sat}")
        
        return iniciales_con_trop, numero_semana
        
    except Exception as e:
        print(f"❌ Error al procesar {archivo_excel}: {e}")
        return [], None

def actualizar_historial_csv(iniciales_con_trop, numero_semana, archivo_csv):
    """Actualiza el archivo CSV con la semana para las personas que tuvieron TROP."""
    try:
        carpeta_destino = os.path.dirname(archivo_csv)
        os.makedirs(carpeta_destino, exist_ok=True)
        
        # Crear archivo CSV si no existe
        if not os.path.exists(archivo_csv):
            df_historial = pd.DataFrame({'empleado': [], 'ultima_semana_trop_sabado': []})
            df_historial.to_csv(archivo_csv, index=False)
            print(f"📄 Archivo CSV creado: {archivo_csv}")
        
        # Leer el archivo CSV actual
        df_historial = pd.read_csv(archivo_csv)
        print(f"\n📊 Archivo CSV actual cargado con {len(df_historial)} registros")
        print("📋 Contenido actual:")
        print(df_historial.to_string(index=False))
        
        # Obtener iniciales encontradas
        iniciales_encontradas = [item['inicial'] for item in iniciales_con_trop]
        print(f"\n🔍 Iniciales con TROP encontradas: {iniciales_encontradas}")
        
        # Actualizar registros existentes y agregar nuevos
        actualizaciones = 0
        nuevas_entradas = 0
        
        for inicial in iniciales_encontradas:
            # Buscar si ya existe el empleado
            empleado_existente = df_historial[df_historial['empleado'] == inicial]
            
            if not empleado_existente.index.empty:
                # Actualizar semana existente
                idx = empleado_existente.index[0]
                valor_anterior = df_historial.at[idx, 'ultima_semana_trop_sabado']
                df_historial.at[idx, 'ultima_semana_trop_sabado'] = numero_semana
                actualizaciones += 1
                print(f"  🔄 Actualizado: {inicial} -> Semana {numero_semana} (antes: {valor_anterior})")
            else:
                # Agregar nuevo empleado
                nueva_fila = {'empleado': inicial, 'ultima_semana_trop_sabado': numero_semana}
                df_historial = pd.concat([df_historial, pd.DataFrame([nueva_fila])], ignore_index=True)
                nuevas_entradas += 1
                print(f"  ➕ Nuevo empleado agregado: {inicial} -> Semana {numero_semana}")
        
        # Convertir la columna a enteros (manteniendo NaN para valores vacíos)
        df_historial['ultima_semana_trop_sabado'] = pd.to_numeric(df_historial['ultima_semana_trop_sabado'], errors='coerce')
        
        # Guardar archivo actualizado con números enteros
        df_historial.to_csv(archivo_csv, index=False, float_format='%.0f')
        
        print(f"\n✅ Archivo CSV actualizado: {actualizaciones} actualizaciones, {nuevas_entradas} nuevas entradas")
        print("📋 Contenido actualizado:")
        print(df_historial.to_string(index=False))
        
        return df_historial
        
    except Exception as e:
        print(f"❌ Error al actualizar CSV: {e}")
        return None

def main():
    print("=== PRUEBA CON ARCHIVO SEMANA 28 ===")
    print("=" * 50)
    
    # Procesar el archivo de la semana 28
    archivo_excel = "horario_descansos_semana_28_1407_2007_2025.xlsx"
    archivo_csv = r"C:\Users\Usuario1\Desktop\cursor\excel_extract\excel_extraction_forschedule\historial_sabados.csv"
    
    if os.path.exists(archivo_excel):
        print(f"📁 Archivo encontrado: {archivo_excel}")
        
        # Extraer datos del Excel
        iniciales_con_trop, numero_semana = extraer_iniciales_con_trop(archivo_excel)
        
        if iniciales_con_trop and numero_semana:
            print(f"\n{'='*50}")
            print(f"📊 RESUMEN DE EXTRACCIÓN:")
            print(f"   • Semana: {numero_semana}")
            print(f"   • Empleados con TROP: {len(iniciales_con_trop)}")
            print(f"   • Iniciales: {', '.join([item['inicial'] for item in iniciales_con_trop])}")
            print(f"{'='*50}")
            
            # Actualizar CSV
            df_actualizado = actualizar_historial_csv(iniciales_con_trop, numero_semana, archivo_csv)
            
            if df_actualizado is not None:
                print(f"\n{'='*50}")
                print(f"✅ PRUEBA COMPLETADA EXITOSAMENTE")
                print(f"{'='*50}")
                print(f"📄 Archivo procesado: {archivo_excel}")
                print(f"📅 Semana: {numero_semana}")
                print(f"👥 Empleados con TROP: {len(iniciales_con_trop)}")
                print(f"📝 Iniciales: {', '.join([item['inicial'] for item in iniciales_con_trop])}")
                print(f"💾 CSV actualizado: {archivo_csv}")
                print(f"{'='*50}\n")
            else:
                print("❌ Error al actualizar el archivo CSV")
        else:
            print("❌ No se encontraron datos válidos para procesar")
    else:
        print(f"❌ Archivo no encontrado: {archivo_excel}")

if __name__ == "__main__":
    main() 