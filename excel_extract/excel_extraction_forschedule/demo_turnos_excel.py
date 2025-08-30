#!/usr/bin/env python3
"""
Demostración del Sistema de Turnos Excel
========================================
Demuestra cómo usar el sistema completo de Excel para TURNOS_FECHAS_ESPECIFICAS
"""

import os
from crear_excel_turnos_especificos import crear_excel_turnos_especificos
from cargar_excel_turnos import cargar_excel_turnos, mostrar_resumen
import json

def demo_completo():
    """
    Demostración completa del sistema
    """
    print("🚀 DEMOSTRACIÓN DEL SISTEMA DE TURNOS EXCEL")
    print("=" * 60)
    
    # Paso 1: Crear el archivo Excel
    print("\n📝 PASO 1: Creando archivo Excel...")
    archivo_excel = crear_excel_turnos_especificos()
    print(f"✅ Archivo creado: {archivo_excel}")
    
    # Paso 2: Verificar que tiene datos de ejemplo
    print("\n📖 PASO 2: Leyendo datos de ejemplo...")
    turnos_ejemplo = cargar_excel_turnos(archivo_excel)
    
    if turnos_ejemplo:
        print("✅ Datos de ejemplo cargados exitosamente")
        mostrar_resumen(turnos_ejemplo)
        
        # Guardar datos de ejemplo
        with open("demo_turnos_ejemplo.json", "w", encoding="utf-8") as f:
            json.dump(turnos_ejemplo, f, ensure_ascii=False, indent=2)
        print("💾 Datos de ejemplo guardados en: demo_turnos_ejemplo.json")
    else:
        print("❌ No se pudieron cargar los datos de ejemplo")
    
    # Paso 3: Mostrar instrucciones
    print("\n📋 PASO 3: INSTRUCCIONES DE USO")
    print("=" * 40)
    print("1. Abra el archivo Excel: TURNOS_FECHAS_ESPECIFICAS.xlsx")
    print("2. Borre las filas de ejemplo (2-5) si desea")
    print("3. Ingrese sus datos:")
    print("   - Columna A: Seleccione empleado del dropdown")
    print("   - Columna B: Seleccione turno del dropdown")
    print("   - Columna C: Fecha inicio (YYYY-MM-DD)")
    print("   - Columna D: Fecha fin (opcional)")
    print("   - Columna E: Comentarios (opcional)")
    print("4. Guarde el archivo Excel")
    print("5. Ejecute: npython cargar_excel_turnos.py")
    print("6. Confirme si desea actualizar config_restricciones.py")
    
    # Paso 4: Mostrar estructura esperada
    print("\n📊 PASO 4: FORMATO DE SALIDA ESPERADO")
    print("=" * 40)
    print("El sistema convierte los datos Excel a:")
    print("TURNOS_FECHAS_ESPECIFICAS = {")
    print('    "JIS": [')
    print('        {"fecha": "2025-07-17", "turno_requerido": "VACA"},')
    print('        {"fecha": "2025-07-18", "turno_requerido": "VACA"},')
    print('        # ... más fechas del rango')
    print('    ],')
    print('    "AFG": [')
    print('        {"fecha": "2025-07-01", "turno_requerido": "COME"},')
    print('        # ... más fechas')
    print('    ]')
    print("}")
    
    print("\n🎯 CARACTERÍSTICAS PRINCIPALES:")
    print("• Dropdown con autocompletado (escriba primera letra)")
    print("• Validación de fechas automática")
    print("• Calendar picker para seleccionar fechas")
    print("• Conversión automática de rangos a fechas individuales")
    print("• Actualización automática del archivo config_restricciones.py")
    print("• Ejemplos incluidos para referencia")
    
    return archivo_excel

def verificar_archivo_excel(archivo):
    """
    Verifica que el archivo Excel existe y es válido
    """
    if not os.path.exists(archivo):
        print(f"❌ El archivo {archivo} no existe")
        return False
    
    try:
        # Intentar cargar el archivo
        turnos = cargar_excel_turnos(archivo)
        print(f"✅ Archivo {archivo} es válido")
        return True
    except Exception as e:
        print(f"❌ Error al verificar el archivo: {e}")
        return False

def main():
    """
    Función principal de demostración
    """
    # Verificar si ya existe el archivo
    archivo_excel = "TURNOS_FECHAS_ESPECIFICAS.xlsx"
    
    if os.path.exists(archivo_excel):
        print(f"📁 El archivo {archivo_excel} ya existe")
        respuesta = input("¿Desea recrearlo? (s/n): ")
        
        if respuesta.lower() in ['s', 'si', 'sí', 'y', 'yes']:
            os.remove(archivo_excel)
            demo_completo()
        else:
            print("✅ Usando archivo existente")
            verificar_archivo_excel(archivo_excel)
    else:
        demo_completo()

if __name__ == "__main__":
    main() 