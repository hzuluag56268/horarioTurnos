#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script para convertir archivos Excel a DataFrames de pandas
Autor: Sistema de Conversión Excel a Pandas
Fecha: 2024
"""

import pandas as pd
import os
import sys
from pathlib import Path


def validar_ruta_archivo(ruta):
    """
    Valida que la ruta ingresada sea válida y el archivo exista.
    
    Args:
        ruta (str): Ruta del archivo a validar
        
    Returns:
        bool: True si la ruta es válida y el archivo existe, False en caso contrario
    """
    try:
        # Convertir a Path para mejor manejo de rutas
        path_archivo = Path(ruta)
        
        # Verificar que la ruta no esté vacía
        if not ruta.strip():
            print("❌ Error: La ruta no puede estar vacía.")
            return False
        
        # Verificar que el archivo existe
        if not path_archivo.exists():
            print(f"❌ Error: El archivo '{ruta}' no existe.")
            return False
        
        # Verificar que es un archivo (no un directorio)
        if not path_archivo.is_file():
            print(f"❌ Error: '{ruta}' no es un archivo válido.")
            return False
        
        # Verificar extensión de archivo Excel
        extensiones_validas = ['.xlsx', '.xls', '.xlsm', '.xlsb']
        if path_archivo.suffix.lower() not in extensiones_validas:
            print(f"❌ Error: El archivo debe tener una extensión Excel válida: {', '.join(extensiones_validas)}")
            return False
        
        return True
        
    except Exception as e:
        print(f"❌ Error al validar la ruta: {str(e)}")
        return False


def cargar_excel_a_dataframe(ruta_archivo):
    """
    Carga un archivo Excel y lo convierte a un DataFrame de pandas.
    
    Args:
        ruta_archivo (str): Ruta del archivo Excel a cargar
        
    Returns:
        pandas.DataFrame: DataFrame cargado del archivo Excel
    """
    try:
        print(f"📂 Cargando archivo: {ruta_archivo}")
        
        # Cargar el archivo Excel
        # engine='openpyxl' para archivos .xlsx, 'xlrd' para archivos .xls
        if Path(ruta_archivo).suffix.lower() == '.xls':
            df = pd.read_excel(ruta_archivo, engine='xlrd')
        else:
            df = pd.read_excel(ruta_archivo, engine='openpyxl')
        
        print("✅ Archivo Excel cargado exitosamente!")
        return df
        
    except FileNotFoundError:
        print(f"❌ Error: No se pudo encontrar el archivo '{ruta_archivo}'")
        raise
    except PermissionError:
        print(f"❌ Error: No tienes permisos para acceder al archivo '{ruta_archivo}'")
        raise
    except Exception as e:
        print(f"❌ Error al cargar el archivo Excel: {str(e)}")
        raise


def mostrar_informacion_dataframe(df):
    """
    Muestra información básica del DataFrame.
    
    Args:
        df (pandas.DataFrame): DataFrame a analizar
    """
    print("\n" + "="*60)
    print("📊 INFORMACIÓN DEL DATAFRAME")
    print("="*60)
    
    # Dimensiones del DataFrame
    print(f"📏 Dimensiones: {df.shape[0]} filas × {df.shape[1]} columnas")
    
    # Información de columnas
    print(f"\n📋 Columnas ({len(df.columns)}):")
    for i, columna in enumerate(df.columns, 1):
        print(f"   {i}. {columna}")
    
    # Tipos de datos
    print(f"\n🔍 Tipos de datos:")
    for columna, tipo in df.dtypes.items():
        print(f"   {columna}: {tipo}")
    
    # Información de valores nulos
    valores_nulos = df.isnull().sum()
    if valores_nulos.sum() > 0:
        print(f"\n⚠️  Valores nulos por columna:")
        for columna, nulos in valores_nulos.items():
            if nulos > 0:
                print(f"   {columna}: {nulos} valores nulos")
    else:
        print(f"\n✅ No hay valores nulos en el DataFrame")
    
    # Primeras filas del DataFrame
    print(f"\n👀 Primeras 5 filas del DataFrame:")
    print("-" * 40)
    print(df.head())
    
    # Últimas filas del DataFrame
    print(f"\n👀 Últimas 5 filas del DataFrame:")
    print("-" * 40)
    print(df.tail())
    
    # Información de memoria
    memoria_mb = df.memory_usage(deep=True).sum() / 1024 / 1024
    print(f"\n💾 Uso de memoria: {memoria_mb:.2f} MB")


def main():
    """
    Función principal del programa.
    """
    print("🚀 CONVERSOR DE EXCEL A DATAFRAME")
    print("="*50)
    print("Este programa convierte archivos Excel a DataFrames de pandas")
    print("="*50)
    
    while True:
        try:
            # Solicitar ruta del archivo al usuario
            print("\n📁 Por favor, ingresa la ruta completa del archivo Excel:")
            print("   Ejemplo: C:\\Users\\Usuario\\Documentos\\archivo.xlsx")
            print("   O: /home/usuario/documentos/archivo.xlsx")
            
            ruta_archivo = input("\n➤ Ruta del archivo: ").strip()
            
            # Validar la ruta ingresada
            if not validar_ruta_archivo(ruta_archivo):
                print("\n🔄 Por favor, intenta nuevamente.")
                continue
            
            # Cargar el archivo Excel
            df = cargar_excel_a_dataframe(ruta_archivo)
            
            # Mostrar información del DataFrame
            mostrar_informacion_dataframe(df)
            
            # Confirmar éxito
            print("\n" + "="*60)
            print("🎉 ¡CONVERSIÓN EXITOSA!")
            print("="*60)
            print(f"✅ El archivo '{os.path.basename(ruta_archivo)}' ha sido convertido")
            print(f"✅ DataFrame creado con {df.shape[0]} filas y {df.shape[1]} columnas")
            print("="*60)
            
            # Preguntar si desea continuar con otro archivo
            print("\n¿Deseas convertir otro archivo? (s/n):")
            continuar = input("➤ ").strip().lower()
            
            if continuar not in ['s', 'si', 'sí', 'y', 'yes']:
                print("\n👋 ¡Gracias por usar el conversor! Hasta luego.")
                break
                
        except KeyboardInterrupt:
            print("\n\n⚠️  Operación cancelada por el usuario.")
            print("👋 ¡Hasta luego!")
            break
        except Exception as e:
            print(f"\n❌ Error inesperado: {str(e)}")
            print("🔄 Por favor, intenta nuevamente.")
            continue


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\n💥 Error crítico del programa: {str(e)}")
        print("🔧 Por favor, verifica que pandas esté instalado correctamente.")
        print("   Puedes instalarlo con: pip install pandas openpyxl xlrd")
        sys.exit(1) 