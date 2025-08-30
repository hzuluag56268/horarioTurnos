#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Funciones simples para convertir Excel ↔ DataFrame
"""

import pandas as pd
from pathlib import Path
from typing import Optional


def excel_to_dataframe(ruta_archivo: str, **kwargs) -> Optional[pd.DataFrame]:
    """
    Convierte un archivo Excel a DataFrame de pandas.
    
    Args:
        ruta_archivo (str): Ruta del archivo Excel
        **kwargs: Argumentos adicionales para pd.read_excel()
        
    Returns:
        pd.DataFrame: DataFrame cargado, None si hay error
    """
    try:
        # Determinar el engine apropiado según la extensión
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


def dataframe_to_excel(df: pd.DataFrame, ruta_salida: str, 
                      sheet_name: str = 'Sheet1', index: bool = False, 
                      **kwargs) -> bool:
    """
    Convierte un DataFrame de pandas a archivo Excel.
    
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