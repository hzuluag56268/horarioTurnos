#!/usr/bin/env python3
"""
Script de Pruebas Extendido: Análisis de Paridad con Tendencias
==============================================================

Versión extendida que analiza más semanas y detecta tendencias temporales.
"""

import sys
import statistics
import io
import contextlib
from generador_descansos_separacion import GeneradorDescansosSeparacion

class AnalizadorParidadExtendido:
    def __init__(self):
        self.resultados_semanas = {}
        self.dias_semana = ['MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT', 'SUN']
        self.nombres_dias = ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado', 'Domingo']
        
    def analizar_rango_extendido(self, semana_inicio, semana_fin, año=2025):
        """Analiza un rango extendido de semanas"""
        print(f"\n🔍 ANÁLISIS EXTENDIDO: SEMANAS {semana_inicio}-{semana_fin} ({año})")
        print("=" * 70)
        
        datos_totales = {dia: [] for dia in self.dias_semana}
        datos_por_semana = {}
        
        for semana in range(semana_inicio, semana_fin + 1):
            print(f"📅 Semana {semana}...", end=" ")
            
            try:
                generador = GeneradorDescansosSeparacion(
                    año=año, mes=1, num_empleados=25, semana_especifica=semana
                )
                
                # Suprimir output
                f = io.StringIO()
                with contextlib.redirect_stdout(f):
                    horario = generador.generar_horario_primera_semana()
                
                # Calcular paridad
                paridad_semana = self._calcular_paridad_semana(horario, generador)
                datos_por_semana[semana] = paridad_semana
                
                # Agregar a datos totales
                for dia in self.dias_semana:
                    if dia in paridad_semana:
                        datos_totales[dia].append(paridad_semana[dia])
                
                print("✅")
                
            except Exception as e:
                print(f"❌")
                continue
        
        # Analizar tendencias
        self._analizar_tendencias(datos_por_semana)
        
        # Estadísticas generales
        self._generar_estadisticas_extendidas(datos_totales)
        
        return datos_totales, datos_por_semana
    
    def _calcular_paridad_semana(self, horario, generador):
        """Calcula paridad diaria para una semana"""
        paridad = {}
        trabajadores_activos = generador._obtener_trabajadores_activos()
        
        for dia_formato in self.dias_semana:
            columna_dia = None
            for col in horario.columns:
                if col.startswith(dia_formato):
                    columna_dia = col
                    break
            
            if columna_dia:
                trabajadores_trabajando = 0
                for idx, empleado in enumerate(generador.empleados):
                    if empleado in trabajadores_activos:
                        turno = horario.iloc[idx][columna_dia]
                        if turno != 'DESC' and turno is not None:
                            trabajadores_trabajando += 1
                
                paridad[dia_formato] = trabajadores_trabajando
        
        return paridad
    
    def _analizar_tendencias(self, datos_por_semana):
        """Analiza tendencias temporales"""
        print(f"\n📈 ANÁLISIS DE TENDENCIAS TEMPORALES")
        print("=" * 70)
        
        semanas_ordenadas = sorted(datos_por_semana.keys())
        
        for i, dia in enumerate(self.dias_semana):
            valores_dia = []
            for semana in semanas_ordenadas:
                if dia in datos_por_semana[semana]:
                    valores_dia.append(datos_por_semana[semana][dia])
            
            if len(valores_dia) >= 3:
                # Calcular tendencia simple (diferencia entre primera y última)
                tendencia = valores_dia[-1] - valores_dia[0]
                tendencia_pct = (tendencia / valores_dia[0]) * 100 if valores_dia[0] > 0 else 0
                
                # Determinar estabilidad
                variaciones = [abs(valores_dia[j] - valores_dia[j-1]) for j in range(1, len(valores_dia))]
                estabilidad = statistics.mean(variaciones) if variaciones else 0
                
                print(f"{self.nombres_dias[i]:<12}: Tendencia {tendencia:+.1f} ({tendencia_pct:+.1f}%), "
                      f"Estabilidad: {estabilidad:.1f}")
    
    def _generar_estadisticas_extendidas(self, datos_totales):
        """Genera estadísticas extendidas"""
        print(f"\n📊 ESTADÍSTICAS EXTENDIDAS")
        print("=" * 70)
        print(f"{'Día':<12} | {'Prom':<6} | {'Min':<3} | {'Max':<3} | {'Rango':<5} | {'CV%':<6} | {'Calidad'}")
        print("-" * 70)
        
        mejores_dias = []
        peores_dias = []
        
        for i, dia in enumerate(self.dias_semana):
            valores = datos_totales[dia]
            
            if valores:
                promedio = statistics.mean(valores)
                minimo = min(valores)
                maximo = max(valores)
                rango = maximo - minimo
                
                if len(valores) > 1:
                    desviacion = statistics.stdev(valores)
                    cv = (desviacion / promedio) * 100 if promedio > 0 else 0
                else:
                    cv = 0
                
                calidad = self._evaluar_calidad(cv)
                nombre_dia = self.nombres_dias[i]
                
                print(f"{nombre_dia:<12} | "
                      f"{promedio:<6.1f} | "
                      f"{minimo:<3} | "
                      f"{maximo:<3} | "
                      f"{rango:<5} | "
                      f"{cv:<6.1f} | "
                      f"{calidad}")
                
                # Clasificar días
                if cv <= 15:
                    mejores_dias.append((nombre_dia, cv))
                elif cv >= 50:
                    peores_dias.append((nombre_dia, cv))
        
        # Resumen de clasificación
        print(f"\n🏆 MEJORES DÍAS (CV ≤ 15%):")
        for dia, cv in sorted(mejores_dias, key=lambda x: x[1]):
            print(f"   ✅ {dia}: {cv:.1f}%")
        
        print(f"\n⚠️ DÍAS PROBLEMÁTICOS (CV ≥ 50%):")
        for dia, cv in sorted(peores_dias, key=lambda x: x[1], reverse=True):
            print(f"   ❌ {dia}: {cv:.1f}%")
    
    def _evaluar_calidad(self, cv):
        """Evalúa calidad basada en CV"""
        if cv <= 10:
            return "✅ Excelente"
        elif cv <= 20:
            return "✅ Buena"
        elif cv <= 30:
            return "⚠️ Regular"
        else:
            return "❌ Mala"
    
    def generar_reporte_comparativo(self, datos1, datos2, nombre1, nombre2):
        """Compara dos conjuntos de datos"""
        print(f"\n🔄 COMPARACIÓN: {nombre1} vs {nombre2}")
        print("=" * 70)
        print(f"{'Día':<12} | {nombre1:<8} | {nombre2:<8} | {'Diferencia':<10} | {'Mejora'}")
        print("-" * 70)
        
        for i, dia in enumerate(self.dias_semana):
            if dia in datos1 and dia in datos2:
                prom1 = statistics.mean(datos1[dia]) if datos1[dia] else 0
                prom2 = statistics.mean(datos2[dia]) if datos2[dia] else 0
                diferencia = prom2 - prom1
                mejora = "✅ Sí" if diferencia > 0 else "❌ No" if diferencia < 0 else "➖ Igual"
                
                nombre_dia = self.nombres_dias[i]
                print(f"{nombre_dia:<12} | "
                      f"{prom1:<8.1f} | "
                      f"{prom2:<8.1f} | "
                      f"{diferencia:<+10.1f} | "
                      f"{mejora}")

def main():
    """Función principal del análisis extendido"""
    print("🧪 ANÁLISIS DE PARIDAD EXTENDIDO")
    print("=" * 50)
    
    analizador = AnalizadorParidadExtendido()
    
    # Análisis de 10 semanas consecutivas
    print("📋 Analizando 10 semanas consecutivas (28-37)...")
    datos_extendidos, datos_por_semana = analizador.analizar_rango_extendido(28, 37)
    
    # Análisis adicional: comparar primeras 5 vs últimas 5 semanas
    datos_primeras = {dia: datos_extendidos[dia][:5] for dia in analizador.dias_semana}
    datos_ultimas = {dia: datos_extendidos[dia][5:] for dia in analizador.dias_semana}
    
    analizador.generar_reporte_comparativo(
        datos_primeras, datos_ultimas, 
        "Sem 28-32", "Sem 33-37"
    )
    
    print(f"\n🎉 ANÁLISIS EXTENDIDO COMPLETADO")
    print("=" * 50)

if __name__ == "__main__":
    main() 