import pandas as pd
import calendar
from datetime import datetime, date
import random
import numpy as np

class GeneradorDescansosParidad:
    def __init__(self, año=2024, mes=7, num_empleados=10):
        self.año = año
        self.mes = mes
        self.num_empleados = num_empleados
        self.empleados = self._generar_empleados()
        self.dias_mes = self._generar_dias_mes()
        self.semanas = self._agrupar_por_semanas()
        
    def _generar_empleados(self):
        """Genera las siglas de los empleados"""
        siglas_base = ['PHD', 'HLG', 'MEI', 'VCM', 'ROP', 'ECE', 'WEH', 'DFB', 'MLS', 'FCE']
        return siglas_base[:self.num_empleados]
    
    def _generar_dias_mes(self):
        """Genera la lista de días del mes con formato DIA-DD"""
        dias = []
        cal = calendar.monthcalendar(self.año, self.mes)
        
        nombres_dias = ['MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT', 'SUN']
        
        for semana in cal:
            for i, dia in enumerate(semana):
                if dia != 0:
                    nombre_dia = nombres_dias[i]
                    formato_dia = f"{nombre_dia}-{dia:02d}"
                    dias.append({
                        'fecha': date(self.año, self.mes, dia),
                        'formato': formato_dia,
                        'dia_semana': i,
                        'es_domingo': i == 6
                    })
        
        return dias
    
    def _agrupar_por_semanas(self):
        """Agrupa los días por semana"""
        semanas = {}
        for dia_info in self.dias_mes:
            semana_num = dia_info['fecha'].isocalendar()[1]
            if semana_num not in semanas:
                semanas[semana_num] = []
            semanas[semana_num].append(dia_info)
        return semanas
    
    def _calcular_paridad_objetivo(self):
        """Calcula el número objetivo de personas descansando por día"""
        total_descansos = self.num_empleados * len(self.semanas) * 2  # 2 descansos por semana
        total_dias = len(self.dias_mes)
        
        # Excluir domingos del cálculo
        dias_no_domingo = [d for d in self.dias_mes if not d['es_domingo']]
        total_dias_disponibles = len(dias_no_domingo)
        
        descansos_por_dia = total_descansos / total_dias_disponibles
        descansos_por_dia_entero = int(descansos_por_dia)
        descansos_extra = total_descansos - (descansos_por_dia_entero * total_dias_disponibles)
        
        print(f"Total descansos a distribuir: {total_descansos}")
        print(f"Días disponibles (sin domingo): {total_dias_disponibles}")
        print(f"Descansos por día objetivo: {descansos_por_dia:.2f}")
        print(f"Descansos por día base: {descansos_por_dia_entero}")
        print(f"Días con descanso extra: {descansos_extra}")
        
        return descansos_por_dia_entero, descansos_extra, dias_no_domingo
    
    def generar_horario_con_paridad(self):
        """Genera el horario con distribución equitativa de descansos"""
        descansos_base, descansos_extra, dias_disponibles = self._calcular_paridad_objetivo()
        
        # Crear DataFrame base
        columnas = ['No.', 'SIGLA ATCO'] + [dia['formato'] for dia in self.dias_mes]
        df = pd.DataFrame(columns=columnas)
        
        # Inicializar contador de descansos por día
        descansos_por_dia = {}
        for dia in dias_disponibles:
            descansos_por_dia[dia['formato']] = 0
        
        # Asignar descansos extra a los primeros días
        dias_con_extra = descansos_extra
        for dia in dias_disponibles:
            if dias_con_extra > 0:
                descansos_por_dia[dia['formato']] = descansos_base + 1
                dias_con_extra -= 1
            else:
                descansos_por_dia[dia['formato']] = descansos_base
        
        # Asignar empleados
        for i, empleado in enumerate(self.empleados):
            fila = {'No.': i + 1, 'SIGLA ATCO': empleado}
            
            # Asignar descansos por semana
            descansos_asignados = self._asignar_descansos_empleado_paridad(i, descansos_por_dia)
            
            # Llenar todos los días
            for dia_info in self.dias_mes:
                formato_dia = dia_info['formato']
                if formato_dia in descansos_asignados:
                    fila[formato_dia] = descansos_asignados[formato_dia]
                else:
                    fila[formato_dia] = None
            
            df = pd.concat([df, pd.DataFrame([fila])], ignore_index=True)
        
        return df
    
    def _asignar_descansos_empleado_paridad(self, empleado_idx, descansos_por_dia):
        """Asigna descansos respetando la paridad diaria"""
        descansos = {}
        
        # Para cada semana
        for semana_num, dias_semana in self.semanas.items():
            dias_semana_disponibles = [d for d in dias_semana if not d['es_domingo']]
            
            if len(dias_semana_disponibles) >= 2:
                # Ordenar días por disponibilidad (menos descansos asignados primero)
                dias_ordenados = sorted(dias_semana_disponibles, 
                                      key=lambda d: descansos_por_dia.get(d['formato'], 0))
                
                # Asignar DESC al día con menos descansos
                if dias_ordenados:
                    descansos[dias_ordenados[0]['formato']] = 'DESC'
                    descansos_por_dia[dias_ordenados[0]['formato']] += 1
                
                # Asignar TROP al siguiente día disponible
                if len(dias_ordenados) > 1:
                    descansos[dias_ordenados[1]['formato']] = 'TROP'
                    descansos_por_dia[dias_ordenados[1]['formato']] += 1
        
        return descansos
    
    def exportar_excel(self, df, nombre_archivo='horario_descansos_paridad_julio.xlsx'):
        """Exporta el horario a Excel"""
        with pd.ExcelWriter(nombre_archivo, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Horario Descansos', index=False)
            
            workbook = writer.book
            worksheet = writer.sheets['Horario Descansos']
            
            # Ajustar ancho de columnas
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 15)
                worksheet.column_dimensions[column_letter].width = adjusted_width
        
        print(f"Horario exportado a: {nombre_archivo}")
        return nombre_archivo
    
    def analizar_paridad(self, df):
        """Analiza la distribución de descansos por día"""
        print("\n=== ANÁLISIS DE PARIDAD DE DESCANSO ===")
        
        # Contar descansos por día
        descansos_por_dia = {}
        for col in df.columns:
            if col.startswith(('MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT', 'SUN')):
                descansos = sum(1 for valor in df[col] if valor in ['DESC', 'TROP'])
                descansos_por_dia[col] = descansos
        
        print("Descansos por día:")
        for dia, count in sorted(descansos_por_dia.items()):
            print(f"  {dia}: {count} personas descansando")
        
        # Estadísticas
        valores = list(descansos_por_dia.values())
        print(f"\nEstadísticas de distribución:")
        print(f"  Promedio: {np.mean(valores):.2f}")
        print(f"  Desviación estándar: {np.std(valores):.2f}")
        print(f"  Mínimo: {min(valores)}")
        print(f"  Máximo: {max(valores)}")
        print(f"  Rango: {max(valores) - min(valores)}")

def main():
    print("=== GENERADOR DE DESCANSO CON PARIDAD - JULIO 2024 ===")
    
    # Crear generador
    generador = GeneradorDescansosParidad(año=2024, mes=7, num_empleados=10)
    
    print(f"Empleados: {generador.empleados}")
    print(f"Días del mes: {len(generador.dias_mes)}")
    
    # Generar horario con paridad
    print("\nGenerando horario con paridad de descansos...")
    horario = generador.generar_horario_con_paridad()
    
    # Analizar paridad
    generador.analizar_paridad(horario)
    
    # Mostrar resumen por empleado
    print("\n=== RESUMEN POR EMPLEADO ===")
    for idx, empleado in enumerate(generador.empleados):
        desc_count = sum(1 for col in horario.columns if col.startswith(('MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT', 'SUN')) 
                        and horario.iloc[idx][col] == 'DESC')
        trop_count = sum(1 for col in horario.columns if col.startswith(('MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT', 'SUN')) 
                        and horario.iloc[idx][col] == 'TROP')
        print(f"Empleado {idx+1} ({empleado}): DESC={desc_count}, TROP={trop_count}")
    
    # Exportar a Excel
    archivo = generador.exportar_excel(horario)
    
    print(f"\n¡Horario con paridad generado exitosamente!")
    print(f"Archivo: {archivo}")
    
    return horario

if __name__ == "__main__":
    horario = main() 