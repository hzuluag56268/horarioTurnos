#!/usr/bin/env python3
"""
Cargador de Excel para Turnos Específicos
==========================================
Lee el archivo Excel y convierte los datos al formato TURNOS_FECHAS_ESPECIFICAS
"""

import openpyxl
import pandas as pd
from datetime import datetime, timedelta
import json
from collections import defaultdict

def cargar_excel_turnos(archivo_excel="TURNOS_FECHAS_ESPECIFICAS.xlsx"):
    """
    Carga los datos del archivo Excel y los convierte a formato TURNOS_FECHAS_ESPECIFICAS
    
    Args:
        archivo_excel (str): Ruta del archivo Excel
        
    Returns:
        dict: Diccionario en formato TURNOS_FECHAS_ESPECIFICAS
    """
    try:
        # Leer el archivo Excel
        df = pd.read_excel(archivo_excel, sheet_name="Turnos Específicos")
        
        # Filtrar filas vacías (donde no hay empleado)
        df = df.dropna(subset=['Empleado'])
        
        # Inicializar diccionario resultado
        turnos_fechas_especificas = defaultdict(list)
        
        # Procesar cada fila
        for index, row in df.iterrows():
            empleado = row['Empleado']
            turno = row['Turno Requerido']
            fecha_inicio = row['Fecha Inicio']
            fecha_fin = row['Fecha Fin']
            
            # Validar datos obligatorios
            if pd.isna(empleado) or pd.isna(turno) or pd.isna(fecha_inicio):
                continue
            
            # Convertir fechas a string
            if isinstance(fecha_inicio, datetime):
                fecha_inicio_str = fecha_inicio.strftime("%Y-%m-%d")
            else:
                fecha_inicio_str = str(fecha_inicio)
            
            # Si no hay fecha fin, usar solo fecha inicio
            if pd.isna(fecha_fin):
                turnos_fechas_especificas[empleado].append({
                    "fecha": fecha_inicio_str,
                    "turno_requerido": turno
                })
            else:
                # Convertir fecha fin
                if isinstance(fecha_fin, datetime):
                    fecha_fin_str = fecha_fin.strftime("%Y-%m-%d")
                else:
                    fecha_fin_str = str(fecha_fin)
                
                # Generar todas las fechas del rango
                fecha_actual = datetime.strptime(fecha_inicio_str, "%Y-%m-%d")
                fecha_final = datetime.strptime(fecha_fin_str, "%Y-%m-%d")
                
                while fecha_actual <= fecha_final:
                    turnos_fechas_especificas[empleado].append({
                        "fecha": fecha_actual.strftime("%Y-%m-%d"),
                        "turno_requerido": turno
                    })
                    fecha_actual += timedelta(days=1)
        
        # Convertir defaultdict a dict normal
        resultado = dict(turnos_fechas_especificas)
        
        return resultado
        
    except Exception as e:
        print(f"❌ Error al cargar el archivo Excel: {e}")
        return {}

def actualizar_config_restricciones(nuevos_turnos, archivo_config="config_restricciones.py"):
    """
    Actualiza el archivo config_restricciones.py con los nuevos turnos
    
    Args:
        nuevos_turnos (dict): Diccionario con los nuevos turnos
        archivo_config (str): Ruta del archivo de configuración
    """
    try:
        # Leer el archivo actual
        with open(archivo_config, 'r', encoding='utf-8') as f:
            contenido = f.read()
        
        # Encontrar el inicio y fin de TURNOS_FECHAS_ESPECIFICAS
        inicio_marker = "TURNOS_FECHAS_ESPECIFICAS = {"
        fin_marker = "}"
        
        inicio_pos = contenido.find(inicio_marker)
        if inicio_pos == -1:
            print("❌ No se encontró TURNOS_FECHAS_ESPECIFICAS en el archivo")
            return False
        
        # Encontrar el cierre del diccionario
        contador_llaves = 0
        pos_actual = inicio_pos + len(inicio_marker)
        
        for i, char in enumerate(contenido[pos_actual:], pos_actual):
            if char == '{':
                contador_llaves += 1
            elif char == '}':
                contador_llaves -= 1
                if contador_llaves == -1:  # Encontramos el cierre
                    fin_pos = i + 1
                    break
        else:
            print("❌ No se pudo encontrar el cierre del diccionario")
            return False
        
        # Generar nuevo contenido del diccionario
        nuevo_dict_str = "TURNOS_FECHAS_ESPECIFICAS = " + dict_to_python_string(nuevos_turnos)
        
        # Reemplazar en el contenido
        nuevo_contenido = contenido[:inicio_pos] + nuevo_dict_str + contenido[fin_pos:]
        
        # Escribir el archivo actualizado
        with open(archivo_config, 'w', encoding='utf-8') as f:
            f.write(nuevo_contenido)
        
        print(f"✅ Archivo {archivo_config} actualizado exitosamente")
        return True
        
    except Exception as e:
        print(f"❌ Error al actualizar el archivo de configuración: {e}")
        return False

def dict_to_python_string(data, indent=0):
    """
    Convierte un diccionario a string con formato Python legible
    """
    if not data:
        return "{}"
    
    result = "{\n"
    for key, value in data.items():
        result += "    " * (indent + 1) + f'"{key}": [\n'
        for item in value:
            result += "    " * (indent + 2) + f'{{"fecha": "{item["fecha"]}", "turno_requerido": "{item["turno_requerido"]}"}},\n'
        result += "    " * (indent + 1) + "],\n"
    result += "    " * indent + "}"
    return result

def mostrar_resumen(turnos_data):
    """
    Muestra un resumen de los turnos cargados
    """
    if not turnos_data:
        print("❌ No hay datos para mostrar")
        return
    
    print("\n📊 RESUMEN DE TURNOS CARGADOS:")
    print("=" * 50)
    
    total_empleados = len(turnos_data)
    total_turnos = sum(len(turnos) for turnos in turnos_data.values())
    
    print(f"👥 Total de empleados: {total_empleados}")
    print(f"📅 Total de turnos específicos: {total_turnos}")
    print()
    
    for empleado, turnos in turnos_data.items():
        print(f"🔹 {empleado}: {len(turnos)} turnos")
        
        # Agrupar por tipo de turno
        turnos_por_tipo = defaultdict(int)
        for turno in turnos:
            turnos_por_tipo[turno["turno_requerido"]] += 1
        
        for tipo, cantidad in turnos_por_tipo.items():
            print(f"   └─ {tipo}: {cantidad} días")
        print()

def main():
    """
    Función principal para cargar y procesar el archivo Excel
    """
    print("🔄 Cargando datos del archivo Excel...")
    
    # Cargar datos del Excel
    turnos_data = cargar_excel_turnos()
    
    if not turnos_data:
        print("❌ No se pudieron cargar los datos")
        return
    
    # Mostrar resumen
    mostrar_resumen(turnos_data)
    
    # Preguntar si actualizar el archivo de configuración
    respuesta = input("\n¿Desea actualizar el archivo config_restricciones.py? (s/n): ")
    
    if respuesta.lower() in ['s', 'si', 'sí', 'y', 'yes']:
        actualizar_config_restricciones(turnos_data)
    else:
        print("✅ Los datos se cargaron correctamente pero no se actualizó el archivo de configuración")
    
    # Guardar como JSON para referencia
    with open("turnos_cargados.json", "w", encoding="utf-8") as f:
        json.dump(turnos_data, f, ensure_ascii=False, indent=2)
    
    print(f"💾 Datos guardados en: turnos_cargados.json")

if __name__ == "__main__":
    main() 