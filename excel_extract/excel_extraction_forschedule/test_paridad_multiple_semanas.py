#!/usr/bin/env python3
"""
Script de Pruebas: Análisis de Paridad Diaria en Múltiples Semanas
================================================================

Analiza la distribución de trabajadores (paridad diaria) a lo largo de 
múltiples semanas en grupos pequeños para evaluar la calidad del algoritmo.

Características:
- Análisis por grupos de 2-4 semanas
- Métricas estadísticas detalladas
- Detección de problemas de paridad
- Reporte consolidado sin generar archivos Excel
"""

import sys
import statistics
import io
import contextlib
from generador_descansos_separacion import GeneradorDescansosSeparacion

class AnalizadorParidadMultiple:
    def __init__(self):
        self.resultados_semanas = {}
        self.dias_semana = ['MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT', 'SUN']
        self.nombres_dias = ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado', 'Domingo']
        
    def analizar_grupo_semanas(self, semana_inicio, semana_fin, año=2025):
        """Analiza un grupo de semanas consecutivas"""
        print(f"\n🔍 ANALIZANDO SEMANAS {semana_inicio}-{semana_fin} ({año})")
        print("=" * 60)
        
        datos_grupo = {dia: [] for dia in self.dias_semana}
        
        for semana in range(semana_inicio, semana_fin + 1):
            print(f"📅 Procesando semana {semana}...", end=" ")
            
            try:
                # Crear generador para esta semana específica
                generador = GeneradorDescansosSeparacion(
                    año=año, 
                    mes=1, 
                    num_empleados=25, 
                    semana_especifica=semana
                )
                
                # Generar horario sin mostrar output detallado
                f = io.StringIO()
                with contextlib.redirect_stdout(f):
                    horario = generador.generar_horario_primera_semana()
                
                # Analizar paridad diaria de esta semana
                paridad_semana = self._calcular_paridad_semana(horario, generador)
                
                # Agregar datos al grupo
                for dia in self.dias_semana:
                    if dia in paridad_semana:
                        datos_grupo[dia].append(paridad_semana[dia])
                
                print("✅")
                
            except Exception as e:
                print(f"❌ Error: {str(e)[:50]}...")
                continue
        
        # Calcular estadísticas del grupo
        estadisticas_grupo = self._calcular_estadisticas_grupo(datos_grupo)
        
        # Guardar resultados
        self.resultados_semanas[f"{semana_inicio}-{semana_fin}"] = {
            'datos': datos_grupo,
            'estadisticas': estadisticas_grupo
        }
        
        return estadisticas_grupo
    
    def _calcular_paridad_semana(self, horario, generador):
        """Calcula la paridad diaria para una semana específica"""
        paridad = {}
        
        # Obtener trabajadores activos (excluir fuera de operación)
        trabajadores_activos = generador._obtener_trabajadores_activos()
        
        # Para cada día de la semana
        for dia_formato in self.dias_semana:
            # Buscar columna correspondiente en el horario
            columna_dia = None
            for col in horario.columns:
                if col.startswith(dia_formato):
                    columna_dia = col
                    break
            
            if columna_dia:
                # Contar trabajadores que NO están de descanso (están trabajando)
                trabajadores_trabajando = 0
                
                for idx, empleado in enumerate(generador.empleados):
                    # Solo contar empleados activos
                    if empleado in trabajadores_activos:
                        turno = horario.iloc[idx][columna_dia]
                        # Si no tiene DESC, está trabajando
                        if turno != 'DESC' and turno is not None:
                            trabajadores_trabajando += 1
                
                paridad[dia_formato] = trabajadores_trabajando
        
        return paridad
    
    def _calcular_estadisticas_grupo(self, datos_grupo):
        """Calcula estadísticas para un grupo de semanas"""
        estadisticas = {}
        
        for dia in self.dias_semana:
            valores = datos_grupo[dia]
            
            if valores:
                promedio = statistics.mean(valores)
                minimo = min(valores)
                maximo = max(valores)
                
                if len(valores) > 1:
                    desviacion = statistics.stdev(valores)
                    coef_variacion = (desviacion / promedio) * 100 if promedio > 0 else 0
                else:
                    desviacion = 0
                    coef_variacion = 0
                
                # Evaluar calidad
                calidad = self._evaluar_calidad_paridad(coef_variacion)
                
                estadisticas[dia] = {
                    'promedio': promedio,
                    'minimo': minimo,
                    'maximo': maximo,
                    'desviacion': desviacion,
                    'coef_variacion': coef_variacion,
                    'calidad': calidad,
                    'valores': valores
                }
        
        return estadisticas
    
    def _evaluar_calidad_paridad(self, coef_variacion):
        """Evalúa la calidad de la paridad basada en coeficiente de variación"""
        if coef_variacion <= 10:
            return "✅ Excelente"
        elif coef_variacion <= 20:
            return "✅ Buena"
        elif coef_variacion <= 30:
            return "⚠️ Regular"
        else:
            return "❌ Mala"
    
    def generar_reporte_grupo(self, grupo_nombre, estadisticas):
        """Genera reporte detallado para un grupo de semanas"""
        print(f"\n📊 REPORTE DE PARIDAD - SEMANAS {grupo_nombre}")
        print("=" * 80)
        print(f"{'Día':<12} | {'Prom':<5} | {'Min':<3} | {'Max':<3} | {'Desv':<5} | {'CV%':<6} | Calidad")
        print("-" * 80)
        
        for i, dia in enumerate(self.dias_semana):
            if dia in estadisticas:
                est = estadisticas[dia]
                nombre_dia = self.nombres_dias[i]
                
                print(f"{nombre_dia:<12} | "
                      f"{est['promedio']:<5.1f} | "
                      f"{est['minimo']:<3} | "
                      f"{est['maximo']:<3} | "
                      f"{est['desviacion']:<5.1f} | "
                      f"{est['coef_variacion']:<6.1f} | "
                      f"{est['calidad']}")
        
        # Resumen de calidad general
        calidades = [est['calidad'] for est in estadisticas.values()]
        excelentes = sum(1 for c in calidades if "Excelente" in c)
        buenas = sum(1 for c in calidades if "Buena" in c)
        regulares = sum(1 for c in calidades if "Regular" in c)
        malas = sum(1 for c in calidades if "Mala" in c)
        
        print(f"\n📈 RESUMEN DE CALIDAD:")
        print(f"   ✅ Excelente: {excelentes} días")
        print(f"   ✅ Buena: {buenas} días") 
        print(f"   ⚠️ Regular: {regulares} días")
        print(f"   ❌ Mala: {malas} días")
    
    def generar_reporte_consolidado(self):
        """Genera reporte consolidado de todos los grupos analizados"""
        print(f"\n🎯 REPORTE CONSOLIDADO - ANÁLISIS COMPLETO")
        print("=" * 80)
        
        if not self.resultados_semanas:
            print("❌ No hay datos para generar reporte")
            return
        
        # Calcular promedios generales por día
        promedios_generales = {}
        
        for dia in self.dias_semana:
            todos_valores = []
            for grupo_datos in self.resultados_semanas.values():
                valores_dia = grupo_datos['datos'][dia]
                todos_valores.extend(valores_dia)
            
            if todos_valores:
                promedio_general = statistics.mean(todos_valores)
                desv_general = statistics.stdev(todos_valores) if len(todos_valores) > 1 else 0
                cv_general = (desv_general / promedio_general) * 100 if promedio_general > 0 else 0
                
                promedios_generales[dia] = {
                    'promedio': promedio_general,
                    'desviacion': desv_general,
                    'coef_variacion': cv_general,
                    'total_muestras': len(todos_valores)
                }
        
        # Mostrar resumen general
        print(f"{'Día':<12} | {'Prom Gral':<9} | {'Desv Gral':<9} | {'CV% Gral':<9} | {'Muestras':<8}")
        print("-" * 80)
        
        for i, dia in enumerate(self.dias_semana):
            if dia in promedios_generales:
                pg = promedios_generales[dia]
                nombre_dia = self.nombres_dias[i]
                
                print(f"{nombre_dia:<12} | "
                      f"{pg['promedio']:<9.1f} | "
                      f"{pg['desviacion']:<9.1f} | "
                      f"{pg['coef_variacion']:<9.1f} | "
                      f"{pg['total_muestras']:<8}")
        
        # Identificar mejor y peor día
        if promedios_generales:
            mejor_dia = min(promedios_generales.items(), key=lambda x: x[1]['coef_variacion'])
            peor_dia = max(promedios_generales.items(), key=lambda x: x[1]['coef_variacion'])
            
            idx_mejor = self.dias_semana.index(mejor_dia[0])
            idx_peor = self.dias_semana.index(peor_dia[0])
            
            print(f"\n🏆 MEJOR DÍA: {self.nombres_dias[idx_mejor]} (CV: {mejor_dia[1]['coef_variacion']:.1f}%)")
            print(f"⚠️ PEOR DÍA: {self.nombres_dias[idx_peor]} (CV: {peor_dia[1]['coef_variacion']:.1f}%)")

def main():
    """Función principal del script de pruebas"""
    print("🧪 SCRIPT DE PRUEBAS: ANÁLISIS DE PARIDAD MÚLTIPLE")
    print("=" * 60)
    print("Analizando distribución de trabajadores en múltiples semanas...")
    
    analizador = AnalizadorParidadMultiple()
    
    # Configuración de pruebas
    grupos_semanas = [
        (28, 29),  # Grupo 1: 2 semanas
        (30, 31),  # Grupo 2: 2 semanas  
        (32, 33),  # Grupo 3: 2 semanas
        (34, 35),  # Grupo 4: 2 semanas
    ]
    
    print(f"\n📋 CONFIGURACIÓN DE PRUEBAS:")
    print(f"   - Grupos a analizar: {len(grupos_semanas)}")
    print(f"   - Total de semanas: {sum(fin - inicio + 1 for inicio, fin in grupos_semanas)}")
    print(f"   - Año: 2025")
    
    # Analizar cada grupo
    for semana_inicio, semana_fin in grupos_semanas:
        estadisticas = analizador.analizar_grupo_semanas(semana_inicio, semana_fin)
        analizador.generar_reporte_grupo(f"{semana_inicio}-{semana_fin}", estadisticas)
    
    # Generar reporte consolidado
    analizador.generar_reporte_consolidado()
    
    print(f"\n🎉 ANÁLISIS COMPLETADO")
    print("=" * 60)

if __name__ == "__main__":
    main() 