#!/usr/bin/env python3
"""
Creador de Excel con xlsxwriter para mejor compatibilidad
========================================================
Versión que usa xlsxwriter para garantizar compatibilidad con Excel
"""

import xlsxwriter
from datetime import datetime, timedelta
import config_restricciones

def crear_excel_xlsxwriter():
    """
    Crea el archivo Excel usando xlsxwriter para mejor compatibilidad
    """
    filename = "TURNOS_FECHAS_ESPECIFICAS.xlsx"
    
    # Crear el workbook
    workbook = xlsxwriter.Workbook(filename)
    
    # Crear hojas
    worksheet = workbook.add_worksheet('Turnos Específicos')
    empleados_ws = workbook.add_worksheet('Lista_Empleados')
    turnos_ws = workbook.add_worksheet('Lista_Turnos')
    instrucciones_ws = workbook.add_worksheet('Instrucciones')
    
    # Obtener datos de configuración
    empleados = config_restricciones.obtener_empleados()
    turnos = config_restricciones.CONFIGURACION_GENERAL["turnos_validos"]
    
    # Configurar formatos
    header_format = workbook.add_format({
        'bold': True,
        'bg_color': '#366092',
        'font_color': 'white',
        'align': 'center',
        'valign': 'vcenter',
        'border': 1
    })
    
    example_format = workbook.add_format({
        'bg_color': '#E6F3FF',
        'border': 1
    })
    
    date_format = workbook.add_format({
        'num_format': 'yyyy-mm-dd',
        'border': 1
    })
    
    cell_format = workbook.add_format({
        'border': 1
    })
    
    # Configurar anchos de columnas
    worksheet.set_column('A:A', 15)  # Empleado
    worksheet.set_column('B:B', 18)  # Turno
    worksheet.set_column('C:C', 15)  # Fecha Inicio
    worksheet.set_column('D:D', 15)  # Fecha Fin
    worksheet.set_column('E:E', 30)  # Comentarios
    
    # Escribir encabezados
    headers = ['Empleado', 'Turno Requerido', 'Fecha Inicio', 'Fecha Fin', 'Comentarios']
    for col, header in enumerate(headers):
        worksheet.write(0, col, header, header_format)
    
    # Escribir datos de ejemplo
    ejemplos = [
        ['JIS', 'VACA', '2025-07-17', '2025-07-30', 'Vacaciones de verano'],
        ['AFG', 'COME', '2025-07-01', '2025-07-31', 'Comisión todo el mes'],
        ['YIS', 'DESC', '2025-07-16', '', 'Descanso un solo día'],
        ['HLG', 'CMED', '2025-07-18', '', 'Cita médica']
    ]
    
    for row, ejemplo in enumerate(ejemplos, 1):
        for col, valor in enumerate(ejemplo):
            if col in [2, 3] and valor:  # Columnas de fecha
                worksheet.write(row, col, valor, example_format)
            else:
                worksheet.write(row, col, valor, example_format)
    
    # Poblar hojas auxiliares
    empleados_ws.write('A1', 'Empleados', header_format)
    for i, empleado in enumerate(empleados):
        empleados_ws.write(i + 1, 0, empleado)
    
    turnos_ws.write('A1', 'Turnos', header_format)
    for i, turno in enumerate(turnos):
        turnos_ws.write(i + 1, 0, turno)
    
    # Ocultar hojas auxiliares
    empleados_ws.hide()
    turnos_ws.hide()
    
    # === VALIDACIONES DE DATOS ===
    
    # Validación para empleados (columna A)
    worksheet.data_validation('A2:A101', {
        'validate': 'list',
        'source': f'=Lista_Empleados!$A$2:$A${len(empleados) + 1}',
        'dropdown': True,
        'error_message': 'Seleccione un empleado válido de la lista.',
        'error_title': 'Empleado inválido',
        'input_message': 'Seleccione un empleado de la lista desplegable.',
        'input_title': 'Empleado'
    })
    
    # Validación para turnos (columna B)
    worksheet.data_validation('B2:B101', {
        'validate': 'list',
        'source': f'=Lista_Turnos!$A$2:$A${len(turnos) + 1}',
        'dropdown': True,
        'error_message': 'Seleccione un turno válido de la lista.',
        'error_title': 'Turno inválido',
        'input_message': 'Seleccione un turno de la lista desplegable.',
        'input_title': 'Turno'
    })
    
    # Validación para fechas
    fecha_minima = datetime.now()
    fecha_maxima = datetime.now() + timedelta(days=730)
    
    worksheet.data_validation('C2:C101', {
        'validate': 'date',
        'criteria': 'between',
        'minimum': fecha_minima,
        'maximum': fecha_maxima,
        'error_message': f'La fecha debe estar entre {fecha_minima.strftime("%Y-%m-%d")} y {fecha_maxima.strftime("%Y-%m-%d")}.',
        'error_title': 'Fecha inválida',
        'input_message': 'Ingrese una fecha válida.',
        'input_title': 'Fecha Inicio'
    })
    
    worksheet.data_validation('D2:D101', {
        'validate': 'date',
        'criteria': 'between',
        'minimum': fecha_minima,
        'maximum': fecha_maxima,
        'error_message': f'La fecha debe estar entre {fecha_minima.strftime("%Y-%m-%d")} y {fecha_maxima.strftime("%Y-%m-%d")}.',
        'error_title': 'Fecha inválida',
        'input_message': 'Ingrese una fecha válida (opcional).',
        'input_title': 'Fecha Fin'
    })
    
    # Aplicar formato a las celdas restantes
    for row in range(5, 101):
        for col in range(5):
            if col in [2, 3]:  # Columnas de fecha
                worksheet.write(row, col, '', date_format)
            else:
                worksheet.write(row, col, '', cell_format)
    
    # Crear hoja de instrucciones
    instrucciones_ws.set_column('A:A', 60)
    
    instrucciones = [
        'INSTRUCCIONES DE USO:',
        '',
        '1. Seleccione empleado del dropdown en columna A',
        '2. Seleccione turno del dropdown en columna B',
        '3. Ingrese fecha inicio en formato YYYY-MM-DD',
        '4. Fecha fin es opcional (si vacía = solo un día)',
        '5. Guarde el archivo después de llenar',
        '6. Ejecute: python cargar_excel_turnos.py',
        '',
        'EMPLEADOS DISPONIBLES:',
        '',
    ]
    
    # Agregar lista de empleados
    for emp in empleados:
        instrucciones.append(f'  • {emp}')
    
    instrucciones.extend([
        '',
        'TURNOS DISPONIBLES:',
        ''
    ])
    
    # Agregar lista de turnos
    for turno in turnos:
        instrucciones.append(f'  • {turno}')
    
    # Escribir instrucciones
    titulo_format = workbook.add_format({
        'bold': True,
        'bg_color': '#366092',
        'font_color': 'white',
        'font_size': 12
    })
    
    normal_format = workbook.add_format({
        'bg_color': '#F0F8FF'
    })
    
    for i, instruccion in enumerate(instrucciones):
        if 'INSTRUCCIONES' in instruccion or 'EMPLEADOS' in instruccion or 'TURNOS' in instruccion:
            instrucciones_ws.write(i, 0, instruccion, titulo_format)
        else:
            instrucciones_ws.write(i, 0, instruccion, normal_format)
    
    # Cerrar el workbook
    workbook.close()
    
    print(f"✅ Archivo Excel creado exitosamente: {filename}")
    print("\n📋 Características del archivo (xlsxwriter):")
    print("   • ✅ Dropdown funcional para empleados")
    print("   • ✅ Dropdown funcional para turnos")
    print("   • ✅ Validación de fechas")
    print("   • ✅ Formato profesional")
    print("   • ✅ Compatibilidad garantizada con Excel")
    print("   • ✅ Instrucciones incluidas")
    
    return filename

if __name__ == "__main__":
    crear_excel_xlsxwriter()
    print("\n🎯 PRUEBA AHORA:")
    print("1. Abra el archivo TURNOS_FECHAS_ESPECIFICAS.xlsx")
    print("2. Haga clic en celda A6 (primera fila vacía)")
    print("3. Debe ver la flecha de dropdown")
    print("4. Haga clic en la flecha para ver empleados")
    print("5. Haga clic en celda B6 para ver turnos")
    print("\nEsta versión usa xlsxwriter que tiene mejor")
    print("compatibilidad con Excel que openpyxl.") 