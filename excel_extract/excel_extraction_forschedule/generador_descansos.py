import pandas as pd
import calendar
from datetime import datetime, date
import random

class GeneradorDescansos:
    def __init__(self, año=2024, mes=7, num_empleados=10):
        self.año = año
        self.mes = mes
        self.num_empleados = num_empleados
        self.empleados = self._generar_empleados()
        self.dias_mes = self._generar_dias_mes()
        
    def _generar_empleados(self):
        """Genera las siglas de los empleados"""
        # Usar siglas de 3 letras como en el ejemplo
        siglas_base = ['PHD', 'HLG', 'MEI', 'VCM', 'ROP', 'ECE', 'WEH', 'DFB', 'MLS', 'FCE']
        return siglas_base[:self.num_empleados]
    
    def _generar_dias_mes(self):
        """Genera la lista de días del mes con formato DIA-DD"""
        dias = []
        cal = calendar.monthcalendar(self.año, self.mes)
        
        # Nombres de días en inglés (como en el formato requerido)
        nombres_dias = ['MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT', 'SUN']
        
        for semana in cal:
            for i, dia in enumerate(semana):
                if dia != 0:  # Si no es 0 (día fuera del mes)
                    nombre_dia = nombres_dias[i]
                    formato_dia = f"{nombre_dia}-{dia:02d}"
                    dias.append({
                        'fecha': date(self.año, self.mes, dia),
                        'formato': formato_dia,
                        'dia_semana': i,  # 0=Lunes, 6=Domingo
                        'es_domingo': i == 6
                    })
        
        return dias
    
    def _asignar_descansos_empleado(self, empleado_idx):
        """Asigna DESC y TROP para un empleado específico"""
        descansos = {}
        
        # Agrupar días por semana
        semanas = {}
        for dia_info in self.dias_mes:
            semana_num = dia_info['fecha'].isocalendar()[1]
            if semana_num not in semanas:
                semanas[semana_num] = []
            semanas[semana_num].append(dia_info)
        
        # Para cada semana, asignar un DESC y un TROP
        for semana_num, dias_semana in semanas.items():
            # Filtrar días disponibles (excluir domingos)
            dias_disponibles = [d for d in dias_semana if not d['es_domingo']]
            
            if len(dias_disponibles) >= 2:
                # Asignar DESC (primer descanso de la semana)
                desc_idx = random.randint(0, len(dias_disponibles) - 1)
                descansos[dias_disponibles[desc_idx]['formato']] = 'DESC'
                
                # Asignar TROP (segundo descanso de la semana)
                dias_restantes = [d for i, d in enumerate(dias_disponibles) if i != desc_idx]
                if dias_restantes:
                    trop_idx = random.randint(0, len(dias_restantes) - 1)
                    descansos[dias_restantes[trop_idx]['formato']] = 'TROP'
        
        return descansos
    
    def generar_horario_descansos(self):
        """Genera el horario completo de descansos"""
        # Crear DataFrame base
        columnas = ['No.', 'SIGLA ATCO'] + [dia['formato'] for dia in self.dias_mes]
        df = pd.DataFrame(columns=columnas)
        
        # Llenar datos de empleados
        for i, empleado in enumerate(self.empleados, 1):
            fila = {'No.': i, 'SIGLA ATCO': empleado}
            
            # Asignar descansos para este empleado
            descansos_empleado = self._asignar_descansos_empleado(i)
            
            # Llenar todos los días
            for dia_info in self.dias_mes:
                formato_dia = dia_info['formato']
                if formato_dia in descansos_empleado:
                    fila[formato_dia] = descansos_empleado[formato_dia]
                else:
                    fila[formato_dia] = None  # Día de trabajo
            
            df = pd.concat([df, pd.DataFrame([fila])], ignore_index=True)
        
        return df
    
    def exportar_excel(self, df, nombre_archivo='horario_descansos_julio.xlsx'):
        """Exporta el horario a Excel con el formato requerido"""
        with pd.ExcelWriter(nombre_archivo, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Horario Descansos', index=False)
            
            # Obtener el workbook y worksheet para formateo
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

def main():
    print("=== GENERADOR DE DESCANSO - JULIO 2024 ===")
    
    # Crear generador
    generador = GeneradorDescansos(año=2024, mes=7, num_empleados=10)
    
    print(f"Empleados: {generador.empleados}")
    print(f"Días del mes: {len(generador.dias_mes)}")
    print(f"Primeros 5 días: {[d['formato'] for d in generador.dias_mes[:5]]}")
    
    # Generar horario
    print("\nGenerando horario de descansos...")
    horario = generador.generar_horario_descansos()
    
    # Mostrar resumen
    print("\n=== RESUMEN DE DESCANSO ===")
    for idx, empleado in enumerate(generador.empleados):
        desc_count = sum(1 for col in horario.columns if col.startswith(('MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT', 'SUN')) 
                        and horario.iloc[idx][col] == 'DESC')
        trop_count = sum(1 for col in horario.columns if col.startswith(('MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT', 'SUN')) 
                        and horario.iloc[idx][col] == 'TROP')
        print(f"Empleado {idx+1} ({empleado}): DESC={desc_count}, TROP={trop_count}")
    
    # Exportar a Excel
    archivo = generador.exportar_excel(horario)
    
    print(f"\n¡Horario generado exitosamente!")
    print(f"Archivo: {archivo}")
    
    return horario

if __name__ == "__main__":
    horario = main() 