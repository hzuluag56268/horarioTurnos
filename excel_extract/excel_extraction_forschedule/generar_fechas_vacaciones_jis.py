from datetime import date, timedelta

def generar_fechas_vacaciones_jis():
    """Genera todas las fechas del rango de vacaciones de JIS"""
    
    # Fechas de inicio y fin
    fecha_inicio = date(2025, 7, 14)  # 14 de julio de 2025
    fecha_fin = date(2025, 8, 4)      # 4 de agosto de 2025
    
    fechas_vacaciones = []
    fecha_actual = fecha_inicio
    
    while fecha_actual <= fecha_fin:
        fechas_vacaciones.append(fecha_actual.strftime('%Y-%m-%d'))
        fecha_actual += timedelta(days=1)
    
    print("=== FECHAS DE VACACIONES PARA JIS ===")
    print(f"Rango: {fecha_inicio.strftime('%d/%m/%Y')} - {fecha_fin.strftime('%d/%m/%Y')}")
    print(f"Total de días: {len(fechas_vacaciones)}")
    print("\nFechas en formato YYYY-MM-DD:")
    
    for i, fecha in enumerate(fechas_vacaciones, 1):
        print(f"  {i:2d}. {fecha}")
    
    # Generar código para agregar al generador
    print("\n=== CÓDIGO PARA AGREGAR AL GENERADOR ===")
    print('"JIS": [')
    for fecha in fechas_vacaciones:
        print(f'    {{"fecha": "{fecha}", "turno_requerido": "VACA"}},')
    print('],')
    
    return fechas_vacaciones

if __name__ == "__main__":
    fechas = generar_fechas_vacaciones_jis() 