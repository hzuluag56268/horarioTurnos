from generador_descansos_separacion import GeneradorDescansosSeparacion

# Crear instancia del generador
generador = GeneradorDescansosSeparacion()

print("=== VERIFICACIÓN DE DÍAS FESTIVOS ===")
print(f"Semana seleccionada: {generador.semana_seleccionada}")
print(f"Fechas de la semana: {generador.fechas_semana[0].strftime('%d/%m/%Y')} - {generador.fechas_semana[6].strftime('%d/%m/%Y')}")

# Obtener días festivos en la semana
dias_festivos = generador._obtener_dias_festivos_semana()

if dias_festivos:
    print("\n🎉 DÍAS FESTIVOS EN LA SEMANA:")
    for dia_festivo in dias_festivos:
        fecha = dia_festivo['fecha']
        formato = dia_festivo['formato_dia']
        print(f"  {formato} ({fecha.strftime('%d/%m/%Y')}): Día festivo - Sin descansos automáticos")
else:
    print("\n📅 No hay días festivos en esta semana")

print(f"\n📋 DÍAS FESTIVOS CONFIGURADOS PARA 2025:")
for fecha_str in generador.dias_festivos:
    fecha = generador.fechas_semana[0].replace(year=2025, month=1, day=1)  # Solo para mostrar formato
    print(f"  {fecha_str}") 