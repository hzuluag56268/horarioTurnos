#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Ejemplo de uso de las funciones excel_to_dataframe y dataframe_to_excel
"""

from excel_functions import excel_to_dataframe, dataframe_to_excel


def ejemplo_basico():
    """
    Ejemplo básico de uso de las funciones
    """
    print("🔄 EJEMPLO BÁSICO DE USO")
    print("="*40)
    
    # Ruta del archivo Excel
    ruta_archivo = r"C:\Users\Usuario1\Desktop\cursor\excel_extract\excel_extraction_forschedule\FormatodeSalidaRequerido.xlsx"
    
    # Convertir Excel a DataFrame
    print(f"📂 Convirtiendo Excel a DataFrame...")
    df = excel_to_dataframe(ruta_archivo)
    
    if df is not None:
        print(f"✅ Excel convertido exitosamente")
        print(f"   Dimensiones: {df.shape[0]} filas × {df.shape[1]} columnas")
        
        # Convertir DataFrame a Excel
        ruta_salida = "archivo_funciones.xlsx"
        print(f"\n💾 Convirtiendo DataFrame a Excel...")
        
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
    print("\n🔄 EJEMPLO CON PARÁMETROS")
    print("="*40)
    
    # Ruta del archivo Excel
    ruta_archivo = r"C:\Users\Usuario1\Desktop\cursor\excel_extract\excel_extraction_forschedule\FormatodeSalidaRequerido.xlsx"
    
    # Convertir Excel a DataFrame con parámetros específicos
    print(f"📂 Convirtiendo Excel a DataFrame con parámetros...")
    df = excel_to_dataframe(ruta_archivo, sheet_name=0, skiprows=0)
    
    if df is not None:
        print(f"✅ Excel convertido exitosamente")
        
        # Convertir DataFrame a Excel con parámetros específicos
        ruta_salida = "archivo_funciones_parametros.xlsx"
        print(f"\n💾 Convirtiendo DataFrame a Excel con parámetros...")
        
        if dataframe_to_excel(df, ruta_salida, sheet_name="Datos", index=False):
            print(f"✅ DataFrame convertido a Excel: {ruta_salida}")
        else:
            print("❌ Error al convertir DataFrame a Excel")
    else:
        print("❌ Error al convertir Excel a DataFrame")


def main():
    """
    Función principal
    """
    print("🚀 EJEMPLOS DE USO DE FUNCIONES")
    print("="*50)
    
    try:
        # Ejecutar ejemplos
        ejemplo_basico()
        ejemplo_con_parametros()
        
        print("\n" + "="*50)
        print("🎉 EJEMPLOS COMPLETADOS")
        print("="*50)
        
    except Exception as e:
        print(f"\n❌ Error: {str(e)}")


if __name__ == "__main__":
    main() 