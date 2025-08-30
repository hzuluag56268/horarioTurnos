#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Clase para convertir archivos Excel a DataFrames de pandas y viceversa
Autor: Sistema de Conversión Excel a Pandas
Fecha: 2024
"""

import pandas as pd
import os
import sys
from pathlib import Path
from typing import Optional, Union, Dict, Any
import logging


class ExcelConverter:
    """
    Clase para convertir archivos Excel a DataFrames de pandas y viceversa.
    
    Esta clase proporciona métodos para:
    - Cargar archivos Excel y convertirlos a DataFrames
    - Exportar DataFrames a archivos Excel
    - Validar rutas y formatos de archivo
    - Mostrar información detallada de los DataFrames
    """
    
    def __init__(self, verbose: bool = True):
        """
        Inicializa el convertidor de Excel.
        
        Args:
            verbose (bool): Si es True, muestra mensajes informativos durante las operaciones
        """
        self.verbose = verbose
        self.extensiones_validas = ['.xlsx', '.xls', '.xlsm', '.xlsb']
        self.ultimo_dataframe = None
        self.ultima_ruta = None
        
        # Configurar logging
        if verbose:
            logging.basicConfig(level=logging.INFO, format='%(message)s')
        else:
            logging.basicConfig(level=logging.WARNING)
        
        self.logger = logging.getLogger(__name__)
    
    def validar_ruta_archivo(self, ruta: str, debe_existir: bool = True) -> bool:
        """
        Valida que la ruta ingresada sea válida y el archivo exista.
        
        Args:
            ruta (str): Ruta del archivo a validar
            debe_existir (bool): Si es True, verifica que el archivo exista
            
        Returns:
            bool: True si la ruta es válida, False en caso contrario
        """
        try:
            # Verificar que la ruta no esté vacía
            if not ruta.strip():
                self.logger.error("❌ Error: La ruta no puede estar vacía.")
                return False
            
            # Convertir a Path para mejor manejo de rutas
            path_archivo = Path(ruta)
            
            # Verificar que el archivo existe (si es requerido)
            if debe_existir and not path_archivo.exists():
                self.logger.error(f"❌ Error: El archivo '{ruta}' no existe.")
                return False
            
            # Verificar que es un archivo (no un directorio) si debe existir
            if debe_existir and not path_archivo.is_file():
                self.logger.error(f"❌ Error: '{ruta}' no es un archivo válido.")
                return False
            
            # Verificar extensión de archivo Excel
            if path_archivo.suffix.lower() not in self.extensiones_validas:
                self.logger.error(f"❌ Error: El archivo debe tener una extensión Excel válida: {', '.join(self.extensiones_validas)}")
                return False
            
            return True
            
        except Exception as e:
            self.logger.error(f"❌ Error al validar la ruta: {str(e)}")
            return False
    
    def cargar_excel(self, ruta_archivo: str, **kwargs) -> pd.DataFrame:
        """
        Carga un archivo Excel y lo convierte a un DataFrame de pandas.
        
        Args:
            ruta_archivo (str): Ruta del archivo Excel a cargar
            **kwargs: Argumentos adicionales para pd.read_excel()
            
        Returns:
            pd.DataFrame: DataFrame cargado del archivo Excel
            
        Raises:
            FileNotFoundError: Si el archivo no existe
            PermissionError: Si no hay permisos para acceder al archivo
            Exception: Para otros errores de lectura
        """
        try:
            # Validar la ruta
            if not self.validar_ruta_archivo(ruta_archivo):
                raise ValueError("Ruta de archivo inválida")
            
            if self.verbose:
                self.logger.info(f"📂 Cargando archivo: {ruta_archivo}")
            
            # Determinar el engine apropiado según la extensión
            extension = Path(ruta_archivo).suffix.lower()
            if extension == '.xls':
                engine = 'xlrd'
            else:
                engine = 'openpyxl'
            
            # Cargar el archivo Excel
            df = pd.read_excel(ruta_archivo, engine=engine, **kwargs)
            
            # Guardar referencia al último DataFrame cargado
            self.ultimo_dataframe = df
            self.ultima_ruta = ruta_archivo
            
            if self.verbose:
                self.logger.info("✅ Archivo Excel cargado exitosamente!")
            
            return df
            
        except FileNotFoundError:
            self.logger.error(f"❌ Error: No se pudo encontrar el archivo '{ruta_archivo}'")
            raise
        except PermissionError:
            self.logger.error(f"❌ Error: No tienes permisos para acceder al archivo '{ruta_archivo}'")
            raise
        except Exception as e:
            self.logger.error(f"❌ Error al cargar el archivo Excel: {str(e)}")
            raise
    
    def exportar_excel(self, df: pd.DataFrame, ruta_salida: str, 
                      sheet_name: str = 'Sheet1', index: bool = False, 
                      **kwargs) -> bool:
        """
        Exporta un DataFrame a un archivo Excel.
        
        Args:
            df (pd.DataFrame): DataFrame a exportar
            ruta_salida (str): Ruta donde guardar el archivo Excel
            sheet_name (str): Nombre de la hoja de Excel
            index (bool): Si incluir el índice del DataFrame
            **kwargs: Argumentos adicionales para df.to_excel()
            
        Returns:
            bool: True si la exportación fue exitosa, False en caso contrario
        """
        try:
            # Validar la ruta de salida (no debe existir necesariamente)
            if not self.validar_ruta_archivo(ruta_salida, debe_existir=False):
                return False
            
            # Verificar que el directorio de destino existe
            directorio_destino = Path(ruta_salida).parent
            if not directorio_destino.exists():
                directorio_destino.mkdir(parents=True, exist_ok=True)
                if self.verbose:
                    self.logger.info(f"📁 Directorio creado: {directorio_destino}")
            
            if self.verbose:
                self.logger.info(f"💾 Exportando DataFrame a: {ruta_salida}")
            
            # Exportar el DataFrame
            df.to_excel(ruta_salida, sheet_name=sheet_name, index=index, 
                       engine='openpyxl', **kwargs)
            
            if self.verbose:
                self.logger.info("✅ DataFrame exportado exitosamente!")
            
            return True
            
        except Exception as e:
            self.logger.error(f"❌ Error al exportar el DataFrame: {str(e)}")
            return False
    
    def mostrar_informacion(self, df: Optional[pd.DataFrame] = None) -> None:
        """
        Muestra información detallada del DataFrame.
        
        Args:
            df (pd.DataFrame, optional): DataFrame a analizar. Si es None, usa el último cargado
        """
        if df is None:
            df = self.ultimo_dataframe
            if df is None:
                self.logger.error("❌ No hay DataFrame disponible para mostrar información.")
                return
        
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
    
    def obtener_estadisticas(self, df: Optional[pd.DataFrame] = None) -> Dict[str, Any]:
        """
        Obtiene estadísticas básicas del DataFrame.
        
        Args:
            df (pd.DataFrame, optional): DataFrame a analizar. Si es None, usa el último cargado
            
        Returns:
            Dict[str, Any]: Diccionario con estadísticas del DataFrame
        """
        if df is None:
            df = self.ultimo_dataframe
            if df is None:
                return {}
        
        return {
            'dimensiones': df.shape,
            'columnas': list(df.columns),
            'tipos_datos': df.dtypes.to_dict(),
            'valores_nulos': df.isnull().sum().to_dict(),
            'memoria_mb': df.memory_usage(deep=True).sum() / 1024 / 1024,
            'columnas_numericas': df.select_dtypes(include=['number']).columns.tolist(),
            'columnas_categoricas': df.select_dtypes(include=['object']).columns.tolist()
        }
    
    def limpiar_dataframe(self, df: Optional[pd.DataFrame] = None, 
                         eliminar_duplicados: bool = True,
                         eliminar_columnas_vacias: bool = True) -> pd.DataFrame:
        """
        Realiza limpieza básica del DataFrame.
        
        Args:
            df (pd.DataFrame, optional): DataFrame a limpiar. Si es None, usa el último cargado
            eliminar_duplicados (bool): Si eliminar filas duplicadas
            eliminar_columnas_vacias (bool): Si eliminar columnas completamente vacías
            
        Returns:
            pd.DataFrame: DataFrame limpio
        """
        if df is None:
            df = self.ultimo_dataframe.copy()
        else:
            df = df.copy()
        
        if self.verbose:
            self.logger.info("🧹 Iniciando limpieza del DataFrame...")
        
        # Eliminar filas duplicadas
        if eliminar_duplicados:
            filas_antes = len(df)
            df = df.drop_duplicates()
            filas_despues = len(df)
            if self.verbose and filas_antes != filas_despues:
                self.logger.info(f"🗑️  Eliminadas {filas_antes - filas_despues} filas duplicadas")
        
        # Eliminar columnas completamente vacías
        if eliminar_columnas_vacias:
            columnas_antes = len(df.columns)
            df = df.dropna(axis=1, how='all')
            columnas_despues = len(df.columns)
            if self.verbose and columnas_antes != columnas_despues:
                self.logger.info(f"🗑️  Eliminadas {columnas_antes - columnas_despues} columnas vacías")
        
        if self.verbose:
            self.logger.info("✅ Limpieza completada")
        
        return df
    
    def convertir_excel_a_dataframe(self, ruta_entrada: str, 
                                   limpiar: bool = True, **kwargs) -> pd.DataFrame:
        """
        Método de conveniencia que carga un Excel y opcionalmente lo limpia.
        
        Args:
            ruta_entrada (str): Ruta del archivo Excel
            limpiar (bool): Si aplicar limpieza automática
            **kwargs: Argumentos adicionales para cargar_excel()
            
        Returns:
            pd.DataFrame: DataFrame procesado
        """
        # Cargar el archivo
        df = self.cargar_excel(ruta_entrada, **kwargs)
        
        # Aplicar limpieza si se solicita
        if limpiar:
            df = self.limpiar_dataframe(df)
            self.ultimo_dataframe = df
        
        return df
    
    def convertir_dataframe_a_excel(self, df: pd.DataFrame, ruta_salida: str,
                                   mostrar_info: bool = True, **kwargs) -> bool:
        """
        Método de conveniencia que exporta un DataFrame a Excel.
        
        Args:
            df (pd.DataFrame): DataFrame a exportar
            ruta_salida (str): Ruta de salida
            mostrar_info (bool): Si mostrar información del DataFrame antes de exportar
            **kwargs: Argumentos adicionales para exportar_excel()
            
        Returns:
            bool: True si la exportación fue exitosa
        """
        if mostrar_info:
            self.mostrar_informacion(df)
        
        return self.exportar_excel(df, ruta_salida, **kwargs)


def main():
    """
    Función principal para demostrar el uso de la clase ExcelConverter.
    """
    print("🚀 CONVERSOR DE EXCEL A DATAFRAME (CLASE)")
    print("="*50)
    
    # Crear instancia del convertidor
    converter = ExcelConverter(verbose=True)
    
    while True:
        try:
            print("\n📁 Por favor, ingresa la ruta completa del archivo Excel:")
            print("   Ejemplo: C:\\Users\\Usuario\\Documentos\\archivo.xlsx")
            
            ruta_archivo = input("\n➤ Ruta del archivo: ").strip()
            
            if not ruta_archivo:
                print("👋 ¡Hasta luego!")
                break
            
            # Cargar y procesar el archivo
            df = converter.convertir_excel_a_dataframe(ruta_archivo, limpiar=True)
            
            # Mostrar información
            converter.mostrar_informacion(df)
            
            # Preguntar si desea exportar
            print("\n¿Deseas exportar el DataFrame a un nuevo archivo Excel? (s/n):")
            exportar = input("➤ ").strip().lower()
            
            if exportar in ['s', 'si', 'sí', 'y', 'yes']:
                print("\n📁 Ingresa la ruta de salida para el nuevo archivo Excel:")
                ruta_salida = input("➤ Ruta de salida: ").strip()
                
                if ruta_salida:
                    exito = converter.convertir_dataframe_a_excel(df, ruta_salida)
                    if exito:
                        print(f"✅ DataFrame exportado exitosamente a: {ruta_salida}")
                    else:
                        print("❌ Error al exportar el DataFrame")
            
            # Preguntar si desea continuar
            print("\n¿Deseas procesar otro archivo? (s/n):")
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