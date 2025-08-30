from datetime import date, timedelta

def calcular_semana_vacaciones():
    """Calcula qué semana del año 2025 incluye el 14 de julio"""
    
    # Calcular el primer lunes de enero 2025
    primer_dia_enero = date(2025, 1, 1)
    dias_hasta_lunes = (7 - primer_dia_enero.weekday()) % 7
    if dias_hasta_lunes == 0:
        dias_hasta_lunes = 7
    primer_lunes_enero = primer_dia_enero + timedelta(days=dias_hasta_lunes)
    
    # Fecha de inicio de vacaciones
    inicio_vacaciones = date(2025, 7, 14)
    
    # Calcular cuántas semanas han pasado desde el primer lunes
    dias_desde_primer_lunes = (inicio_vacaciones - primer_lunes_enero).days
    semana_vacaciones = (dias_desde_primer_lunes // 7) + 1
    
    # Calcular el lunes de esa semana
    lunes_semana = primer_lunes_enero + timedelta(days=(semana_vacaciones - 1) * 7)
    domingo_semana = lunes_semana + timedelta(days=6)
    
    print("=== CÁLCULO DE SEMANA DE VACACIONES JIS ===")
    print(f"Primer lunes de enero 2025: {primer_lunes_enero.strftime('%d/%m/%Y')}")
    print(f"Inicio de vacaciones JIS: {inicio_vacaciones.strftime('%d/%m/%Y')}")
    print(f"Semana número: {semana_vacaciones}")
    print(f"Lunes de esa semana: {lunes_semana.strftime('%d/%m/%Y')}")
    print(f"Domingo de esa semana: {domingo_semana.strftime('%d/%m/%Y')}")
    
    # Verificar si el 14 de julio cae en esa semana
    if lunes_semana <= inicio_vacaciones <= domingo_semana:
        print(f"✅ El 14 de julio SÍ cae en la semana {semana_vacaciones}")
    else:
        print(f"❌ El 14 de julio NO cae en la semana {semana_vacaciones}")
    
    return semana_vacaciones

if __name__ == "__main__":
    semana = calcular_semana_vacaciones() 