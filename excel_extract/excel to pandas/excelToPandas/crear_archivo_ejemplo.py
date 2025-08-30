#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script para crear un archivo Excel de ejemplo para probar el conversor
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import random

def crear_archivo_ejemplo():
    """
    Crea un archivo Excel de ejemplo con datos variados para probar el conversor.
    """
    
    # Generar datos de ejemplo
    np.random.seed(42)  # Para reproducibilidad
    
    # Lista de nombres
    nombres = ['Juan', 'María', 'Pedro', 'Ana', 'Carlos', 'Laura', 'Miguel', 'Sofía', 
               'David', 'Carmen', 'Javier', 'Elena', 'Roberto', 'Isabel', 'Fernando']
    
    # Lista de ciudades
    ciudades = ['Madrid', 'Barcelona', 'Valencia', 'Sevilla', 'Málaga', 'Bilbao', 
                'Zaragoza', 'Murcia', 'Palma', 'Las Palmas', 'Alicante', 'Córdoba']
    
    # Generar datos
    n_registros = 100
    
    datos = {
        'ID': range(1, n_registros + 1),
        'Nombre': [random.choice(nombres) for _ in range(n_registros)],
        'Edad': np.random.randint(18, 65, n_registros),
        'Ciudad': [random.choice(ciudades) for _ in range(n_registros)],
        'Salario': np.random.uniform(25000, 80000, n_registros).round(2),
        'Departamento': [random.choice(['Ventas', 'IT', 'RRHH', 'Finanzas', 'Marketing']) 
                        for _ in range(n_registros)],
        'Fecha_Contratacion': [datetime.now() - timedelta(days=random.randint(1, 1000)) 
                              for _ in range(n_registros)],
        'Activo': [random.choice([True, False]) for _ in range(n_registros)]
    }
    
    # Crear DataFrame
    df = pd.DataFrame(datos)
    
    # Agregar algunos valores nulos para probar el manejo
    df.loc[random.sample(range(n_registros), 5), 'Salario'] = np.nan
    df.loc[random.sample(range(n_registros), 3), 'Ciudad'] = None
    
    # Guardar como Excel
    nombre_archivo = 'datos_ejemplo.xlsx'
    df.to_excel(nombre_archivo, index=False, engine='openpyxl')
    
    print(f"✅ Archivo de ejemplo creado: {nombre_archivo}")
    print(f"📊 Datos generados: {n_registros} registros con 8 columnas")
    print(f"📋 Columnas: {', '.join(df.columns.tolist())}")
    print(f"💾 Tamaño del archivo: {df.memory_usage(deep=True).sum() / 1024:.2f} KB")
    
    # Mostrar vista previa
    print(f"\n👀 Vista previa de los datos:")
    print("-" * 50)
    print(df.head())
    
    print(f"\n🎯 Ahora puedes usar este archivo para probar el conversor:")
    print(f"   python excel_to_dataframe.py")
    print(f"   Y cuando te pregunte la ruta, ingresa: {nombre_archivo}")

if __name__ == "__main__":
    try:
        crear_archivo_ejemplo()
    except Exception as e:
        print(f"❌ Error al crear el archivo de ejemplo: {str(e)}")
        print("🔧 Asegúrate de tener pandas y openpyxl instalados:")
        print("   pip install pandas openpyxl") 