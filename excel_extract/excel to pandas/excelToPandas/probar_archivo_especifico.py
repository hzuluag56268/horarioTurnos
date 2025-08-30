#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script para probar la clase ExcelConverter con el archivo específico
"""

from excel_converter import ExcelConverter
import os


def probar_archivo_especifico():
    """
    Prueba la clase ExcelConverter con el archivo específico mencionado
    """
    print("🚀 PROBANDO CLASE EXCELCONVERTER")
    print("="*50)
    
    # Crear instancia del convertidor
    converter = ExcelConverter(verbose=True)
    
    # Ruta del archivo específico
    ruta_archivo = r"C:\Users\Usuario1\Desktop\cursor\excel_extract\excel_extraction_forschedule\FormatodeSalidaRequerido.xlsx"
    
    print(f"📁 Archivo a procesar: {ruta_archivo}")
    
    try:
        # Verificar que el archivo existe
        if not os.path.exists(ruta_archivo):
            print(f"❌ Error: El archivo no existe en la ruta especificada")
            return
        
        print(f"✅ Archivo encontrado: {os.path.basename(ruta_archivo)}")
        print(f"📏 Tamaño: {os.path.getsize(ruta_archivo) / 1024:.2f} KB")
        
        # Cargar el archivo Excel
        print(f"\n📂 Cargando archivo Excel...")
        df = converter.convertir_excel_a_dataframe(ruta_archivo, limpiar=True)
        
        # Mostrar información del DataFrame
        converter.mostrar_informacion(df)
        
        # Obtener estadísticas
        stats = converter.obtener_estadisticas(df)
        
        print(f"\n📊 RESUMEN DE ESTADÍSTICAS:")
        print(f"   Dimensiones: {stats['dimensiones']}")
        print(f"   Memoria utilizada: {stats['memoria_mb']:.2f} MB")
        print(f"   Columnas numéricas: {len(stats['columnas_numericas'])}")
        print(f"   Columnas categóricas: {len(stats['columnas_categoricas'])}")
        
        # Preguntar si desea exportar
        print(f"\n¿Deseas exportar el DataFrame procesado? (s/n):")
        exportar = input("➤ ").strip().lower()
        
        if exportar in ['s', 'si', 'sí', 'y', 'yes']:
            # Crear nombre de archivo de salida
            nombre_base = os.path.splitext(os.path.basename(ruta_archivo))[0]
            ruta_salida = f"{nombre_base}_procesado.xlsx"
            
            print(f"\n💾 Exportando a: {ruta_salida}")
            
            if converter.convertir_dataframe_a_excel(df, ruta_salida, mostrar_info=False):
                print(f"✅ Archivo exportado exitosamente: {ruta_salida}")
            else:
                print("❌ Error al exportar el archivo")
        
        print(f"\n🎉 Procesamiento completado exitosamente!")
        
    except FileNotFoundError:
        print(f"❌ Error: No se pudo encontrar el archivo '{ruta_archivo}'")
    except PermissionError:
        print(f"❌ Error: No tienes permisos para acceder al archivo '{ruta_archivo}'")
    except Exception as e:
        print(f"❌ Error inesperado: {str(e)}")


def probar_con_archivo_ejemplo():
    """
    Prueba la clase con el archivo de ejemplo generado
    """
    print(f"\n" + "="*50)
    print("🔄 PROBANDO CON ARCHIVO DE EJEMPLO")
    print("="*50)
    
    converter = ExcelConverter(verbose=True)
    
    # Usar el archivo de ejemplo que ya existe
    ruta_ejemplo = "datos_ejemplo.xlsx"
    
    if os.path.exists(ruta_ejemplo):
        print(f"📁 Probando con archivo de ejemplo: {ruta_ejemplo}")
        
        try:
            # Cargar archivo
            df = converter.convertir_excel_a_dataframe(ruta_ejemplo, limpiar=True)
            
            # Mostrar información básica
            print(f"✅ Archivo cargado: {df.shape[0]} filas, {df.shape[1]} columnas")
            
            # Exportar versión procesada
            ruta_salida = "datos_ejemplo_procesado.xlsx"
            if converter.convertir_dataframe_a_excel(df, ruta_salida, mostrar_info=False):
                print(f"✅ Archivo procesado exportado: {ruta_salida}")
            
        except Exception as e:
            print(f"❌ Error: {str(e)}")
    else:
        print(f"⚠️  Archivo de ejemplo no encontrado: {ruta_ejemplo}")


def main():
    """
    Función principal
    """
    print("🔍 PROBANDO CLASE EXCELCONVERTER")
    print("="*60)
    
    # Probar con el archivo específico
    probar_archivo_especifico()
    
    # Probar con archivo de ejemplo
    probar_con_archivo_ejemplo()
    
    print(f"\n" + "="*60)
    print("🎉 PRUEBAS COMPLETADAS")
    print("="*60)


if __name__ == "__main__":
    main() 