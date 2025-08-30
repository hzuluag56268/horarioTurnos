#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script para probar la conversión con la ruta específica proporcionada
"""

from excel_converter_simple import excel_to_dataframe, dataframe_to_excel


def probar_conversion():
    """
    Prueba la conversión con la ruta específica
    """
    print("🚀 PROBANDO CONVERSIÓN CON RUTA ESPECÍFICA")
    print("="*50)
    
    # Ruta específica proporcionada por el usuario
    ruta_archivo = r"C:\Users\Usuario1\Desktop\cursor\excel_extract\excel_extraction_forschedule\FormatodeSalidaRequerido.xlsx"
    
    print(f"📁 Archivo a procesar: {ruta_archivo}")
    
    # Convertir Excel a DataFrame
    print("\n📂 Convirtiendo Excel a DataFrame...")
    df = excel_to_dataframe(ruta_archivo)
    
    if df is not None:
        print(f"✅ Excel convertido a DataFrame exitosamente")
        print(f"   Dimensiones: {df.shape[0]} filas × {df.shape[1]} columnas")
        
        # Convertir DataFrame a Excel
        ruta_salida = "archivo_convertido_desde_input.xlsx"
        print(f"\n💾 Exportando DataFrame a: {ruta_salida}")
        
        if dataframe_to_excel(df, ruta_salida):
            print(f"✅ DataFrame exportado exitosamente a: {ruta_salida}")
            print(f"\n🎉 Conversión completada exitosamente!")
        else:
            print("❌ Error al exportar el DataFrame")
    else:
        print("❌ Error al convertir Excel a DataFrame")
        print("   Verifica que el archivo existe y tiene un formato válido")


if __name__ == "__main__":
    probar_conversion() 