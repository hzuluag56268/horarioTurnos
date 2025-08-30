#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Ejemplo de uso de la clase ExcelConverter
Este script demuestra cómo usar la clase para convertir Excel ↔ DataFrame
"""

from excel_converter import ExcelConverter
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import random


def crear_datos_ejemplo():
    """
    Crea un DataFrame de ejemplo para demostrar las funcionalidades.
    """
    print("📊 Creando datos de ejemplo...")
    
    # Generar datos de ejemplo
    np.random.seed(42)
    
    nombres = ['Juan', 'María', 'Pedro', 'Ana', 'Carlos', 'Laura', 'Miguel', 'Sofía']
    ciudades = ['Madrid', 'Barcelona', 'Valencia', 'Sevilla', 'Málaga', 'Bilbao']
    departamentos = ['Ventas', 'IT', 'RRHH', 'Finanzas', 'Marketing']
    
    n_registros = 50
    
    datos = {
        'ID': range(1, n_registros + 1),
        'Nombre': [random.choice(nombres) for _ in range(n_registros)],
        'Edad': np.random.randint(18, 65, n_registros),
        'Ciudad': [random.choice(ciudades) for _ in range(n_registros)],
        'Salario': np.random.uniform(25000, 80000, n_registros).round(2),
        'Departamento': [random.choice(departamentos) for _ in range(n_registros)],
        'Fecha_Contratacion': [datetime.now() - timedelta(days=random.randint(1, 1000)) 
                              for _ in range(n_registros)],
        'Activo': [random.choice([True, False]) for _ in range(n_registros)]
    }
    
    df = pd.DataFrame(datos)
    
    # Agregar algunos valores nulos para probar limpieza
    df.loc[random.sample(range(n_registros), 3), 'Salario'] = np.nan
    df.loc[random.sample(range(n_registros), 2), 'Ciudad'] = None
    
    return df


def ejemplo_conversion_basica():
    """
    Ejemplo básico de conversión Excel ↔ DataFrame
    """
    print("\n" + "="*60)
    print("🔄 EJEMPLO 1: CONVERSIÓN BÁSICA")
    print("="*60)
    
    # Crear instancia del convertidor
    converter = ExcelConverter(verbose=True)
    
    # Crear datos de ejemplo
    df_ejemplo = crear_datos_ejemplo()
    
    # Exportar DataFrame a Excel
    ruta_salida = "datos_ejemplo_clase.xlsx"
    print(f"\n📤 Exportando DataFrame a: {ruta_salida}")
    
    if converter.convertir_dataframe_a_excel(df_ejemplo, ruta_salida):
        print("✅ Exportación exitosa!")
        
        # Cargar el archivo Excel de vuelta
        print(f"\n📥 Cargando archivo Excel: {ruta_salida}")
        df_cargado = converter.convertir_excel_a_dataframe(ruta_salida, limpiar=True)
        
        # Mostrar información del DataFrame cargado
        converter.mostrar_informacion(df_cargado)
        
        # Verificar que los datos son iguales
        if df_ejemplo.equals(df_cargado):
            print("✅ Los datos se mantuvieron idénticos durante la conversión")
        else:
            print("⚠️  Los datos cambiaron durante la conversión")
    else:
        print("❌ Error en la exportación")


def ejemplo_limpieza_avanzada():
    """
    Ejemplo de limpieza avanzada de datos
    """
    print("\n" + "="*60)
    print("🧹 EJEMPLO 2: LIMPIEZA AVANZADA")
    print("="*60)
    
    converter = ExcelConverter(verbose=True)
    
    # Crear datos con duplicados y valores nulos
    df_sucio = crear_datos_ejemplo()
    
    # Agregar filas duplicadas
    filas_duplicadas = df_sucio.iloc[:3].copy()
    df_sucio = pd.concat([df_sucio, filas_duplicadas], ignore_index=True)
    
    # Agregar columna completamente vacía
    df_sucio['Columna_Vacia'] = None
    
    print(f"📊 DataFrame original: {df_sucio.shape[0]} filas, {df_sucio.shape[1]} columnas")
    print(f"🔍 Valores nulos: {df_sucio.isnull().sum().sum()}")
    
    # Aplicar limpieza
    df_limpio = converter.limpiar_dataframe(df_sucio, 
                                           eliminar_duplicados=True,
                                           eliminar_columnas_vacias=True)
    
    print(f"\n📊 DataFrame limpio: {df_limpio.shape[0]} filas, {df_limpio.shape[1]} columnas")
    print(f"🔍 Valores nulos: {df_limpio.isnull().sum().sum()}")


def ejemplo_estadisticas():
    """
    Ejemplo de obtención de estadísticas
    """
    print("\n" + "="*60)
    print("📈 EJEMPLO 3: ESTADÍSTICAS")
    print("="*60)
    
    converter = ExcelConverter(verbose=True)
    
    # Crear datos de ejemplo
    df = crear_datos_ejemplo()
    
    # Obtener estadísticas
    stats = converter.obtener_estadisticas(df)
    
    print("📊 Estadísticas del DataFrame:")
    print(f"   Dimensiones: {stats['dimensiones']}")
    print(f"   Columnas: {len(stats['columnas'])}")
    print(f"   Memoria: {stats['memoria_mb']:.2f} MB")
    print(f"   Columnas numéricas: {len(stats['columnas_numericas'])}")
    print(f"   Columnas categóricas: {len(stats['columnas_categoricas'])}")
    
    print(f"\n📋 Tipos de datos:")
    for col, tipo in stats['tipos_datos'].items():
        print(f"   {col}: {tipo}")
    
    print(f"\n⚠️  Valores nulos:")
    for col, nulos in stats['valores_nulos'].items():
        if nulos > 0:
            print(f"   {col}: {nulos}")


def ejemplo_uso_avanzado():
    """
    Ejemplo de uso avanzado con múltiples archivos
    """
    print("\n" + "="*60)
    print("🚀 EJEMPLO 4: USO AVANZADO")
    print("="*60)
    
    converter = ExcelConverter(verbose=True)
    
    # Crear múltiples DataFrames
    df1 = crear_datos_ejemplo()
    df2 = crear_datos_ejemplo()
    df2['ID'] = range(51, 101)  # IDs diferentes
    
    # Exportar múltiples hojas a un solo archivo Excel
    ruta_multiple = "datos_multiple_hojas.xlsx"
    
    print(f"📤 Exportando múltiples DataFrames a: {ruta_multiple}")
    
    # Exportar primera hoja
    if converter.exportar_excel(df1, ruta_multiple, sheet_name='Empleados_1'):
        print("✅ Primera hoja exportada")
        
        # Exportar segunda hoja (usando pandas directamente para múltiples hojas)
        with pd.ExcelWriter(ruta_multiple, engine='openpyxl', mode='a') as writer:
            df2.to_excel(writer, sheet_name='Empleados_2', index=False)
        print("✅ Segunda hoja exportada")
        
        # Cargar ambas hojas
        print(f"\n📥 Cargando hojas del archivo: {ruta_multiple}")
        
        df_cargado1 = converter.cargar_excel(ruta_multiple, sheet_name='Empleados_1')
        df_cargado2 = converter.cargar_excel(ruta_multiple, sheet_name='Empleados_2')
        
        print(f"📊 Hoja 1: {df_cargado1.shape[0]} filas")
        print(f"📊 Hoja 2: {df_cargado2.shape[0]} filas")


def ejemplo_manejo_errores():
    """
    Ejemplo de manejo de errores
    """
    print("\n" + "="*60)
    print("🛡️ EJEMPLO 5: MANEJO DE ERRORES")
    print("="*60)
    
    converter = ExcelConverter(verbose=True)
    
    # Intentar cargar archivo que no existe
    print("🔍 Probando carga de archivo inexistente...")
    try:
        df = converter.cargar_excel("archivo_inexistente.xlsx")
    except FileNotFoundError:
        print("✅ Error manejado correctamente: Archivo no encontrado")
    
    # Intentar cargar archivo con extensión inválida
    print("\n🔍 Probando carga de archivo con extensión inválida...")
    if not converter.validar_ruta_archivo("archivo.txt"):
        print("✅ Validación correcta: Extensión no válida")
    
    # Intentar exportar a directorio sin permisos (simulado)
    print("\n🔍 Probando exportación...")
    df_ejemplo = crear_datos_ejemplo()
    if converter.exportar_excel(df_ejemplo, "datos_test.xlsx"):
        print("✅ Exportación exitosa")


def main():
    """
    Función principal que ejecuta todos los ejemplos
    """
    print("🚀 EJEMPLOS DE USO DE LA CLASE EXCELCONVERTER")
    print("="*60)
    
    try:
        # Ejecutar todos los ejemplos
        ejemplo_conversion_basica()
        ejemplo_limpieza_avanzada()
        ejemplo_estadisticas()
        ejemplo_uso_avanzado()
        ejemplo_manejo_errores()
        
        print("\n" + "="*60)
        print("🎉 TODOS LOS EJEMPLOS COMPLETADOS EXITOSAMENTE")
        print("="*60)
        print("📁 Archivos generados:")
        print("   - datos_ejemplo_clase.xlsx")
        print("   - datos_multiple_hojas.xlsx")
        print("   - datos_test.xlsx")
        
    except Exception as e:
        print(f"\n❌ Error durante la ejecución: {str(e)}")


if __name__ == "__main__":
    main() 