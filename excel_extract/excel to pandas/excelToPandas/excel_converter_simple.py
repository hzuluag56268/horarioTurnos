#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Clase simplificada para convertir archivos Excel a DataFrames de pandas y viceversa
Solo conversión, sin mostrar información ni modificar datos
Solicita al usuario la ruta del archivo mediante input
"""

import pandas as pd
from pathlib import Path
from typing import Optional


class ExcelConverterSimple:
    """
    Clase simplificada para convertir archivos Excel a DataFrames de pandas y viceversa.
    Solo realiza conversión sin mostrar información ni modificar datos.
    """
    
    def __init__(self):
        """
        Inicializa el convertidor simple.
        """
        self.extensiones_validas = ['.xlsx', '.xls', '.xlsm', '.xlsb']
    
    def validar_ruta(self, ruta: str, debe_existir: bool = True) -> bool:
        """
        Valida que la ruta sea válida.
        
        Args:
            ruta (str): Ruta del archivo a validar
            debe_existir (bool): Si verificar que el archivo exista
            
        Returns:
            bool: True si la ruta es válida
        """
        try:
            if not ruta.strip():
                return False
            
            path_archivo = Path(ruta)
            
            if debe_existir and not path_archivo.exists():
                return False
            
            if debe_existir and not path_archivo.is_file():
                return False
            
            if path_archivo.suffix.lower() not in self.extensiones_validas:
                return False
            
            return True
            
        except Exception:
            return False
    
    def excel_a_dataframe(self, ruta_archivo: str, **kwargs) -> Optional[pd.DataFrame]:
        """
        Convierte un archivo Excel a DataFrame.
        
        Args:
            ruta_archivo (str): Ruta del archivo Excel
            **kwargs: Argumentos adicionales para pd.read_excel()
            
        Returns:
            pd.DataFrame: DataFrame cargado, None si hay error
        """
        try:
            if not self.validar_ruta(ruta_archivo):
                return None
            
            # Determinar el engine apropiado
            extension = Path(ruta_archivo).suffix.lower()
            if extension == '.xls':
                engine = 'xlrd'
            else:
                engine = 'openpyxl'
            
            # Cargar el archivo Excel
            df = pd.read_excel(ruta_archivo, engine=engine, **kwargs)
            return df
            
        except Exception:
            return None
    
    def dataframe_a_excel(self, df: pd.DataFrame, ruta_salida: str, 
                         sheet_name: str = 'Sheet1', index: bool = False, 
                         **kwargs) -> bool:
        """
        Convierte un DataFrame a archivo Excel.
        
        Args:
            df (pd.DataFrame): DataFrame a exportar
            ruta_salida (str): Ruta donde guardar el archivo Excel
            sheet_name (str): Nombre de la hoja de Excel
            index (bool): Si incluir el índice del DataFrame
            **kwargs: Argumentos adicionales para df.to_excel()
            
        Returns:
            bool: True si la exportación fue exitosa
        """
        try:
            if not self.validar_ruta(ruta_salida, debe_existir=False):
                return False
            
            # Verificar que el directorio de destino existe
            directorio_destino = Path(ruta_salida).parent
            if not directorio_destino.exists():
                directorio_destino.mkdir(parents=True, exist_ok=True)
            
            # Exportar el DataFrame
            df.to_excel(ruta_salida, sheet_name=sheet_name, index=index, 
                       engine='openpyxl', **kwargs)
            return True
            
        except Exception:
            return False


# Funciones de conveniencia para uso directo
def excel_to_dataframe(ruta_archivo: str, **kwargs) -> Optional[pd.DataFrame]:
    """
    Función simple para convertir Excel a DataFrame.
    
    Args:
        ruta_archivo (str): Ruta del archivo Excel
        **kwargs: Argumentos adicionales para pd.read_excel()
        
    Returns:
        pd.DataFrame: DataFrame cargado, None si hay error
    """
    converter = ExcelConverterSimple()
    return converter.excel_a_dataframe(ruta_archivo, **kwargs)


def dataframe_to_excel(df: pd.DataFrame, ruta_salida: str, 
                      sheet_name: str = 'Sheet1', index: bool = False, 
                      **kwargs) -> bool:
    """
    Función simple para convertir DataFrame a Excel.
    
    Args:
        df (pd.DataFrame): DataFrame a exportar
        ruta_salida (str): Ruta donde guardar el archivo Excel
        sheet_name (str): Nombre de la hoja de Excel
        index (bool): Si incluir el índice del DataFrame
        **kwargs: Argumentos adicionales para df.to_excel()
        
    Returns:
        bool: True si la exportación fue exitosa
    """
    converter = ExcelConverterSimple()
    return converter.dataframe_a_excel(df, ruta_salida, sheet_name, index, **kwargs)


def solicitar_ruta_archivo() -> str:
    """
    Solicita al usuario la ruta del archivo Excel mediante input.
    
    Returns:
        str: Ruta del archivo ingresada por el usuario
    """
    print("🚀 CONVERSOR DE EXCEL A DATAFRAME")
    print("="*50)
    print("Este programa convierte archivos Excel a DataFrames de pandas")
    print("="*50)
    
    print("\n📁 Por favor, ingresa la ruta completa del archivo Excel:")
    print("   Ejemplo: C:\\Users\\Usuario\\Documentos\\archivo.xlsx")
    print("   O: /home/usuario/documentos/archivo.xlsx")
    print("   Extensiones válidas: .xlsx, .xls, .xlsm, .xlsb")
    
    while True:
        ruta = input("\n➤ Ruta del archivo: ").strip()
        
        if not ruta:
            print("❌ Error: La ruta no puede estar vacía. Intenta nuevamente.")
            continue
        
        # Validar la ruta
        converter = ExcelConverterSimple()
        if converter.validar_ruta(ruta):
            return ruta
        else:
            print("❌ Error: Ruta inválida o archivo no encontrado. Intenta nuevamente.")


def main():
    """
    Función principal que solicita la ruta y realiza la conversión.
    """
    try:
        # Solicitar ruta del archivo
        ruta_archivo = solicitar_ruta_archivo()
        
        print(f"\n📂 Procesando archivo: {ruta_archivo}")
        
        # Convertir Excel a DataFrame
        df = excel_to_dataframe(ruta_archivo)
        
        if df is not None:
            print(f"✅ Excel convertido a DataFrame exitosamente")
            print(f"   Dimensiones: {df.shape[0]} filas × {df.shape[1]} columnas")
            
            # Preguntar si desea exportar
            print(f"\n¿Deseas exportar el DataFrame a un archivo Excel? (s/n):")
            exportar = input("➤ ").strip().lower()
            
            if exportar in ['s', 'si', 'sí', 'y', 'yes']:
                print(f"\n📁 Ingresa la ruta de salida para el archivo Excel:")
                ruta_salida = input("➤ Ruta de salida: ").strip()
                
                if ruta_salida:
                    if dataframe_to_excel(df, ruta_salida):
                        print(f"✅ DataFrame exportado exitosamente a: {ruta_salida}")
                    else:
                        print("❌ Error al exportar el DataFrame")
                else:
                    print("❌ Ruta de salida no válida")
            
            print(f"\n🎉 Conversión completada exitosamente!")
            
        else:
            print("❌ Error al convertir Excel a DataFrame")
            print("   Verifica que el archivo existe y tiene un formato válido")
        
    except KeyboardInterrupt:
        print("\n\n⚠️  Operación cancelada por el usuario.")
    except Exception as e:
        print(f"\n❌ Error inesperado: {str(e)}")


# Ejemplo de uso directo
if __name__ == "__main__":
    main() 