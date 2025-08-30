#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Ejemplo de uso de la versión simplificada de ExcelConverter
Solo conversión, sin mostrar información ni modificar datos
"""

from excel_converter_simple import ExcelConverterSimple, excel_to_dataframe, dataframe_to_excel


def ejemplo_uso_clase():
    """
    Ejemplo usando la clase ExcelConverterSimple
    """
    print("🔄 Ejemplo usando la clase ExcelConverterSimple")
    
    # Crear instancia
    converter = ExcelConverterSimple()
    
    # Ruta del archivo a convertir
    ruta_archivo = r"C:\Users\Usuario1\Desktop\cursor\excel_extract\excel_extraction_forschedule\FormatodeSalidaRequerido.xlsx"
    
    # Convertir Excel a DataFrame
    df = converter.excel_a_dataframe(ruta_archivo)
    
    if df is not None:
        print(f"✅ Excel convertido a DataFrame exitosamente")
        print(f"   Dimensiones: {df.shape[0]} filas × {df.shape[1]} columnas")
        
        # Convertir DataFrame a Excel
        ruta_salida = "archivo_convertido.xlsx"
        if converter.dataframe_a_excel(df, ruta_salida):
            print(f"✅ DataFrame convertido a Excel: {ruta_salida}")
        else:
            print("❌ Error al convertir DataFrame a Excel")
    else:
        print("❌ Error al convertir Excel a DataFrame")


def ejemplo_uso_funciones():
    """
    Ejemplo usando las funciones de conveniencia
    """
    print("\n🔄 Ejemplo usando funciones de conveniencia")
    
    # Ruta del archivo a convertir
    ruta_archivo = r"C:\Users\Usuario1\Desktop\cursor\excel_extract\excel_extraction_forschedule\FormatodeSalidaRequerido.xlsx"
    
    # Convertir Excel a DataFrame
    df = excel_to_dataframe(ruta_archivo)
    
    if df is not None:
        print(f"✅ Excel convertido a DataFrame exitosamente")
        print(f"   Dimensiones: {df.shape[0]} filas × {df.shape[1]} columnas")
        
        # Convertir DataFrame a Excel
        ruta_salida = "archivo_convertido_funcion.xlsx"
        if dataframe_to_excel(df, ruta_salida):
            print(f"✅ DataFrame convertido a Excel: {ruta_salida}")
        else:
            print("❌ Error al convertir DataFrame a Excel")
    else:
        print("❌ Error al convertir Excel a DataFrame")


def ejemplo_con_parametros():
    """
    Ejemplo usando parámetros adicionales
    """
    print("\n🔄 Ejemplo con parámetros adicionales")
    
    # Ruta del archivo a convertir
    ruta_archivo = r"C:\Users\Usuario1\Desktop\cursor\excel_extract\excel_extraction_forschedule\FormatodeSalidaRequerido.xlsx"
    
    # Convertir Excel a DataFrame con parámetros específicos
    df = excel_to_dataframe(ruta_archivo, sheet_name=0, skiprows=0)
    
    if df is not None:
        print(f"✅ Excel convertido a DataFrame con parámetros específicos")
        
        # Convertir DataFrame a Excel con parámetros específicos
        ruta_salida = "archivo_convertido_parametros.xlsx"
        if dataframe_to_excel(df, ruta_salida, sheet_name="Datos", index=False):
            print(f"✅ DataFrame convertido a Excel con parámetros: {ruta_salida}")
        else:
            print("❌ Error al convertir DataFrame a Excel")
    else:
        print("❌ Error al convertir Excel a DataFrame")


def main():
    """
    Función principal
    """
    print("🚀 EJEMPLOS DE USO SIMPLIFICADO")
    print("="*50)
    
    try:
        # Ejecutar ejemplos
        ejemplo_uso_clase()
        ejemplo_uso_funciones()
        ejemplo_con_parametros()
        
        print("\n" + "="*50)
        print("🎉 TODOS LOS EJEMPLOS COMPLETADOS")
        print("="*50)
        
    except Exception as e:
        print(f"\n❌ Error: {str(e)}")


if __name__ == "__main__":
    main() 