import pandas as pd
import numpy as np

def analizar_solucion_optimizada(archivo='horario_optimizado_or_tools_julio.xlsx'):
    """Analiza la solución optimizada con OR-Tools"""
    try:
        df = pd.read_excel(archivo)
        print(f"✅ Archivo cargado: {archivo}")
    except FileNotFoundError:
        print(f"❌ Archivo no encontrado: {archivo}")
        return
    
    print("\n" + "="*60)
    print("📊 ANÁLISIS DE SOLUCIÓN OPTIMIZADA CON OR-TOOLS")
    print("="*60)
    
    # 1. ANÁLISIS DE SÁBADOS (OBJETIVO PRINCIPAL)
    print("\n🎯 1. MAXIMIZACIÓN DE SÁBADOS")
    print("-" * 40)
    
    sabados_por_empleado = []
    total_descansos_sabados = 0
    
    print("Sábados por empleado:")
    for idx, empleado in enumerate(df['SIGLA ATCO']):
        sabados_count = sum(1 for col in df.columns if col.startswith('SAT') 
                          and df.iloc[idx][col] in ['DESC', 'TROP'])
        sabados_por_empleado.append(sabados_count)
        total_descansos_sabados += sabados_count
        print(f"  {empleado}: {sabados_count} sábados")
    
    print(f"\n📈 Total descansos en sábados: {total_descansos_sabados}")
    print(f"📈 Máximo posible: 8 (4 sábados × 2 descansos)")
    print(f"📈 Porcentaje de aprovechamiento: {(total_descansos_sabados/8)*100:.1f}%")
    
    if total_descansos_sabados >= 6:
        print("✅ EXCELENTE: Máxima utilización de sábados")
    elif total_descansos_sabados >= 4:
        print("✅ BUENO: Buena utilización de sábados")
    else:
        print("⚠️  MEJORABLE: Baja utilización de sábados")
    
    # 2. ANÁLISIS DE PARIDAD DIARIA
    print("\n⚖️ 2. PARIDAD DIARIA")
    print("-" * 40)
    
    descansos_por_dia = {}
    for col in df.columns:
        if col.startswith(('MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT', 'SUN')):
            descansos = sum(1 for valor in df[col] if valor in ['DESC', 'TROP'])
            descansos_por_dia[col] = descansos
    
    print("Descansos por día:")
    for dia, count in sorted(descansos_por_dia.items()):
        print(f"  {dia}: {count} personas descansando")
    
    valores = list(descansos_por_dia.values())
    print(f"\n📊 Estadísticas de paridad:")
    print(f"  Promedio: {np.mean(valores):.2f}")
    print(f"  Desviación estándar: {np.std(valores):.2f}")
    print(f"  Mínimo: {min(valores)}")
    print(f"  Máximo: {max(valores)}")
    print(f"  Rango: {max(valores) - min(valores)}")
    
    if max(valores) - min(valores) <= 2:
        print("✅ EXCELENTE: Paridad diaria muy equilibrada")
    elif max(valores) - min(valores) <= 4:
        print("✅ BUENA: Paridad diaria equilibrada")
    else:
        print("⚠️  MEJORABLE: Paridad diaria desequilibrada")
    
    # 3. ANÁLISIS DE SEPARACIÓN DE DESCANSO
    print("\n📅 3. SEPARACIÓN DE DESCANSO")
    print("-" * 40)
    
    empleados_con_consecutivos = 0
    total_semanas_consecutivas = 0
    
    for idx, empleado in enumerate(df['SIGLA ATCO']):
        print(f"\n{empleado}:")
        consecutivos_empleado = 0
        
        # Agrupar por semanas (aproximadamente)
        semanas = {
            27: ['MON-01', 'TUE-02', 'WED-03', 'THU-04', 'FRI-05', 'SAT-06'],
            28: ['MON-08', 'TUE-09', 'WED-10', 'THU-11', 'FRI-12', 'SAT-13'],
            29: ['MON-15', 'TUE-16', 'WED-17', 'THU-18', 'FRI-19', 'SAT-20'],
            30: ['MON-22', 'TUE-23', 'WED-24', 'THU-25', 'FRI-26', 'SAT-27'],
            31: ['MON-29', 'TUE-30', 'WED-31']
        }
        
        for semana_num, dias_semana in semanas.items():
            descansos_semana = []
            for dia in dias_semana:
                if dia in df.columns:
                    valor = df.iloc[idx][dia]
                    if valor in ['DESC', 'TROP']:
                        # Obtener día de la semana (0=Lunes, 5=Sábado)
                        dia_semana = int(dia.split('-')[0].replace('MON', '0').replace('TUE', '1')
                                       .replace('WED', '2').replace('THU', '3').replace('FRI', '4').replace('SAT', '5'))
                        descansos_semana.append((dia, valor, dia_semana))
            
            if len(descansos_semana) == 2:
                dia1, tipo1, num_dia1 = descansos_semana[0]
                dia2, tipo2, num_dia2 = descansos_semana[1]
                separacion = abs(num_dia1 - num_dia2)
                
                if separacion == 1:
                    print(f"  ⚠️  Semana {semana_num}: {dia1}({tipo1}) y {dia2}({tipo2}) - CONSECUTIVOS!")
                    consecutivos_empleado += 1
                    total_semanas_consecutivas += 1
                elif separacion >= 3:
                    print(f"  ✅ Semana {semana_num}: {dia1}({tipo1}) y {dia2}({tipo2}) - Separación: {separacion} días")
                else:
                    print(f"  ⚠️  Semana {semana_num}: {dia1}({tipo1}) y {dia2}({tipo2}) - Separación mínima: {separacion} días")
        
        if consecutivos_empleado > 0:
            empleados_con_consecutivos += 1
    
    print(f"\n📊 Resumen de separación:")
    print(f"  Empleados con descansos consecutivos: {empleados_con_consecutivos}/{len(df)}")
    print(f"  Total semanas con descansos consecutivos: {total_semanas_consecutivas}")
    
    if empleados_con_consecutivos == 0:
        print("✅ EXCELENTE: Sin descansos consecutivos")
    elif empleados_con_consecutivos <= 2:
        print("✅ BUENO: Pocos descansos consecutivos")
    else:
        print("⚠️  MEJORABLE: Muchos descansos consecutivos")
    
    # 4. RESUMEN GENERAL
    print("\n🏆 4. RESUMEN GENERAL")
    print("-" * 40)
    
    print(f"📈 Sábados aprovechados: {total_descansos_sabados}/8 ({(total_descansos_sabados/8)*100:.1f}%)")
    print(f"⚖️  Paridad diaria: Rango {max(valores) - min(valores)} (objetivo: ≤2)")
    print(f"📅 Separación: {empleados_con_consecutivos} empleados con consecutivos")
    
    # Calcular puntuación general
    puntuacion_sabados = min(100, (total_descansos_sabados/8)*100)
    puntuacion_paridad = max(0, 100 - (max(valores) - min(valores))*20)
    puntuacion_separacion = max(0, 100 - empleados_con_consecutivos*10)
    
    puntuacion_total = (puntuacion_sabados + puntuacion_paridad + puntuacion_separacion) / 3
    
    print(f"\n🎯 PUNTUACIÓN GENERAL: {puntuacion_total:.1f}/100")
    
    if puntuacion_total >= 90:
        print("🏆 EXCELENTE: Solución muy optimizada")
    elif puntuacion_total >= 80:
        print("🥇 MUY BUENA: Solución bien optimizada")
    elif puntuacion_total >= 70:
        print("🥈 BUENA: Solución aceptable")
    else:
        print("🥉 MEJORABLE: Necesita optimización")

if __name__ == "__main__":
    analizar_solucion_optimizada() 