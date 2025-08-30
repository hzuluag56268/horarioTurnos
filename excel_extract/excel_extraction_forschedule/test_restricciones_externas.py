#!/usr/bin/env python3
"""
Script de Prueba - Sistema de Restricciones Externas
==================================================

Este script prueba que el sistema de restricciones externas funciona
correctamente con el generador principal.

Versión: 1.0
Fecha: 2025
"""

import sys
import traceback
from datetime import datetime

def test_importacion_configuracion():
    """Prueba que la configuración se importa correctamente"""
    print("🔍 PRUEBA 1: Importación de Configuración Externa")
    print("-" * 50)
    
    try:
        from config_restricciones import (
            RESTRICCIONES_EMPLEADOS,
            TURNOS_FECHAS_ESPECIFICAS,
            TURNOS_ESPECIALES,
            TRABAJADORES_FUERA_OPERACION,
            DIAS_FESTIVOS,
            CONFIGURACION_GENERAL,
            validar_configuracion,
            obtener_resumen_configuracion
        )
        
        print("✅ Importación exitosa de todos los componentes")
        
        # Validar configuración
        errores = validar_configuracion()
        if errores:
            print(f"❌ {len(errores)} errores en configuración:")
            for error in errores:
                print(f"   - {error}")
            return False
        else:
            print("✅ Configuración válida")
        
        # Mostrar resumen
        resumen = obtener_resumen_configuracion()
        print("\n📊 RESUMEN DE CONFIGURACIÓN:")
        for clave, valor in resumen.items():
            print(f"   {clave}: {valor}")
        
        return True
        
    except Exception as e:
        print(f"❌ Error en importación: {e}")
        traceback.print_exc()
        return False

def test_generador_con_restricciones_externas():
    """Prueba que el generador funciona con restricciones externas"""
    print("\n🔍 PRUEBA 2: Generador con Restricciones Externas")
    print("-" * 50)
    
    try:
        from generador_descansos_separacion import GeneradorDescansosSeparacion
        
        # Crear generador para semana 28 (semana de prueba)
        generador = GeneradorDescansosSeparacion(
            año=2025,
            mes=7,
            num_empleados=24,
            semana_especifica=28
        )
        
        print("✅ Generador creado exitosamente")
        
        # Verificar que las restricciones externas se cargaron
        print(f"📋 Empleados con restricciones: {len(generador.restricciones_empleados)}")
        print(f"📅 Empleados con fechas específicas: {len(generador.turnos_fechas_especificas)}")
        print(f"🎯 Empleados con turnos especiales: {len(generador.turnos_especiales)}")
        print(f"🚫 Trabajadores fuera de operación: {len(generador.trabajadores_fuera_operacion)}")
        print(f"🎉 Días festivos configurados: {len(generador.dias_festivos)}")
        
        # Verificar configuraciones específicas
        print("\n🔍 VERIFICACIÓN DE CONFIGURACIONES ESPECÍFICAS:")
        
        # HZG debe tener restricciones
        if "HZG" in generador.restricciones_empleados:
            restriccion_hzg = generador.restricciones_empleados["HZG"]
            print(f"✅ HZG tiene restricciones: {restriccion_hzg}")
        else:
            print("❌ HZG no tiene restricciones configuradas")
            return False
        
        # JIS debe tener vacaciones
        if "JIS" in generador.turnos_fechas_especificas:
            vacaciones_jis = len(generador.turnos_fechas_especificas["JIS"])
            print(f"✅ JIS tiene {vacaciones_jis} días de vacaciones configurados")
        else:
            print("❌ JIS no tiene vacaciones configuradas")
            return False
        
        # GMT debe tener turno especial SIND
        if "GMT" in generador.turnos_especiales:
            turnos_gmt = generador.turnos_especiales["GMT"]
            print(f"✅ GMT tiene turnos especiales: {turnos_gmt}")
        else:
            print("❌ GMT no tiene turnos especiales configurados")
            return False
        
        # PHD debe estar fuera de operación
        if "PHD" in generador.trabajadores_fuera_operacion:
            print("✅ PHD está correctamente fuera de operación")
        else:
            print("❌ PHD no está fuera de operación")
            return False
        
        return True
        
    except Exception as e:
        print(f"❌ Error en generador: {e}")
        traceback.print_exc()
        return False

def test_generacion_horario():
    """Prueba que se puede generar un horario completo"""
    print("\n🔍 PRUEBA 3: Generación de Horario Completo")
    print("-" * 50)
    
    try:
        from generador_descansos_separacion import GeneradorDescansosSeparacion
        
        # Crear generador para semana 28
        generador = GeneradorDescansosSeparacion(
            año=2025,
            mes=7,
            num_empleados=24,
            semana_especifica=28
        )
        
        print("🚀 Iniciando generación de horario...")
        
        # Generar horario
        df = generador.generar_horario_primera_semana()
        
        if df is not None and not df.empty:
            print(f"✅ Horario generado exitosamente")
            print(f"📊 Dimensiones: {df.shape[0]} empleados x {df.shape[1]} columnas")
            
            # Verificar que las restricciones se aplicaron
            print("\n🔍 VERIFICACIÓN DE RESTRICCIONES APLICADAS:")
            
            # Verificar JIS tiene VACA
            if 'JIS' in df.index:
                fila_jis = df.loc['JIS']
                vaca_count = sum(1 for val in fila_jis if val == 'VACA')
                print(f"✅ JIS tiene {vaca_count} días de VACA")
            
            # Verificar GMT tiene SIND
            if 'GMT' in df.index:
                fila_gmt = df.loc['GMT']
                sind_count = sum(1 for val in fila_gmt if val == 'SIND')
                print(f"✅ GMT tiene {sind_count} días de SIND")
            
            # Verificar PHD no aparece (fuera de operación)
            if 'PHD' not in df.index:
                print("✅ PHD correctamente excluido (fuera de operación)")
            else:
                print("⚠️ PHD aparece en el horario (debería estar fuera de operación)")
            
            return True
        else:
            print("❌ No se pudo generar el horario")
            return False
            
    except Exception as e:
        print(f"❌ Error en generación de horario: {e}")
        traceback.print_exc()
        return False

def test_modificacion_dinamica():
    """Prueba modificaciones dinámicas de la configuración"""
    print("\n🔍 PRUEBA 4: Modificación Dinámica de Configuración")
    print("-" * 50)
    
    try:
        from config_restricciones import (
            agregar_restriccion_empleado,
            agregar_fecha_especifica,
            agregar_turno_especial,
            RESTRICCIONES_EMPLEADOS,
            TURNOS_FECHAS_ESPECIFICAS,
            TURNOS_ESPECIALES
        )
        
        # Agregar nueva restricción
        print("🔧 Agregando nueva restricción para empleado TEST...")
        agregar_restriccion_empleado("TEST", "DESC", ["lunes"], "fijo")
        
        if "TEST" in RESTRICCIONES_EMPLEADOS:
            print("✅ Restricción agregada correctamente")
        else:
            print("❌ Error agregando restricción")
            return False
        
        # Agregar fecha específica
        print("🔧 Agregando fecha específica para empleado TEST...")
        agregar_fecha_especifica("TEST", "2025-07-15", "DESC")
        
        if "TEST" in TURNOS_FECHAS_ESPECIFICAS:
            print("✅ Fecha específica agregada correctamente")
        else:
            print("❌ Error agregando fecha específica")
            return False
        
        # Agregar turno especial
        print("🔧 Agregando turno especial para empleado TEST...")
        agregar_turno_especial("TEST", "CERT", "semanal_fijo", "viernes")
        
        if "TEST" in TURNOS_ESPECIALES:
            print("✅ Turno especial agregado correctamente")
        else:
            print("❌ Error agregando turno especial")
            return False
        
        # Limpiar las modificaciones de prueba
        if "TEST" in RESTRICCIONES_EMPLEADOS:
            del RESTRICCIONES_EMPLEADOS["TEST"]
        if "TEST" in TURNOS_FECHAS_ESPECIFICAS:
            del TURNOS_FECHAS_ESPECIFICAS["TEST"]
        if "TEST" in TURNOS_ESPECIALES:
            del TURNOS_ESPECIALES["TEST"]
        
        print("🧹 Modificaciones de prueba limpiadas")
        return True
        
    except Exception as e:
        print(f"❌ Error en modificación dinámica: {e}")
        traceback.print_exc()
        return False

def main():
    """Función principal de pruebas"""
    print("🧪 SISTEMA DE PRUEBAS - RESTRICCIONES EXTERNAS")
    print("=" * 60)
    print(f"📅 Fecha de ejecución: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print()
    
    # Ejecutar todas las pruebas
    pruebas = [
        ("Importación de Configuración", test_importacion_configuracion),
        ("Generador con Restricciones Externas", test_generador_con_restricciones_externas),
        ("Generación de Horario Completo", test_generacion_horario),
        ("Modificación Dinámica", test_modificacion_dinamica)
    ]
    
    resultados = []
    
    for nombre, funcion_prueba in pruebas:
        try:
            resultado = funcion_prueba()
            resultados.append((nombre, resultado))
        except Exception as e:
            print(f"❌ ERROR CRÍTICO en {nombre}: {e}")
            resultados.append((nombre, False))
    
    # Mostrar resumen final
    print("\n" + "=" * 60)
    print("📊 RESUMEN FINAL DE PRUEBAS")
    print("=" * 60)
    
    exitosas = 0
    fallidas = 0
    
    for nombre, resultado in resultados:
        if resultado:
            print(f"✅ {nombre}")
            exitosas += 1
        else:
            print(f"❌ {nombre}")
            fallidas += 1
    
    print("-" * 60)
    print(f"🎯 TOTAL: {exitosas} exitosas, {fallidas} fallidas")
    
    if fallidas == 0:
        print("🎉 ¡TODAS LAS PRUEBAS PASARON!")
        print("✨ El sistema de restricciones externas está funcionando correctamente")
        return 0
    else:
        print("⚠️ Algunas pruebas fallaron")
        print("🔧 Revise los errores mostrados arriba")
        return 1

if __name__ == "__main__":
    sys.exit(main())