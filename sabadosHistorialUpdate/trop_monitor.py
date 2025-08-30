import pandas as pd
import openpyxl
import re
import os
import time
import shutil
from datetime import datetime
from pathlib import Path
import logging

# Configurar logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('trop_monitor.log'),
        logging.StreamHandler()
    ]
)

class TropMonitor:
    def __init__(self):
        # Rutas de las carpetas
        self.carpeta_monitoreo = r"C:\Users\Usuario1\Desktop\cursor\sabadosHistorialUpdate"
        self.carpeta_destino = r"C:\Users\Usuario1\Desktop\cursor\excel_extract\excel_extraction_forschedule"
        self.archivo_csv = os.path.join(self.carpeta_destino, "historial_sabados.csv")
        
        # Crear carpetas si no existen
        os.makedirs(self.carpeta_monitoreo, exist_ok=True)
        os.makedirs(self.carpeta_destino, exist_ok=True)
        
        # Archivos procesados para evitar reprocesamiento
        self.archivos_procesados = set()
        
        logging.info(f"Carpeta de monitoreo: {self.carpeta_monitoreo}")
        logging.info(f"Carpeta de destino: {self.carpeta_destino}")
        logging.info(f"Archivo CSV: {self.archivo_csv}")

    def extraer_numero_semana(self, nombre_archivo):
        """Extrae el número de semana del nombre del archivo Excel."""
        match = re.search(r'semana_(\d+)', nombre_archivo)
        if match:
            return int(match.group(1))
        return None

    def extraer_iniciales_con_trop(self, archivo_excel):
        """Extrae las iniciales de empleados que tienen 'TROP' en columnas SAT."""
        try:
            logging.info(f"Procesando archivo: {archivo_excel}")
            
            # Extraer número de semana
            numero_semana = self.extraer_numero_semana(archivo_excel)
            logging.info(f"Número de semana detectado: {numero_semana}")
            
            # Leer el archivo Excel
            excel_file = pd.ExcelFile(archivo_excel)
            hojas = excel_file.sheet_names
            logging.info(f"Hojas encontradas: {hojas}")
            
            iniciales_con_trop = []
            
            for hoja in hojas:
                logging.info(f"Procesando hoja: {hoja}")
                
                df = pd.read_excel(archivo_excel, sheet_name=hoja)
                
                # Buscar columnas que empiecen con 'SAT'
                columnas_sat = [col for col in df.columns if str(col).startswith('SAT')]
                logging.info(f"Columnas SAT encontradas: {columnas_sat}")
                
                if columnas_sat:
                    for col_sat in columnas_sat:
                        # Buscar filas con 'TROP'
                        trop_filas = df[df[col_sat].astype(str).str.contains('TROP', case=False, na=False)]
                        
                        if not trop_filas.empty:
                            logging.info(f"Encontradas {len(trop_filas)} filas con TROP en {col_sat}")
                            
                            for idx, fila in trop_filas.iterrows():
                                inicial = None
                                
                                # Buscar iniciales en las primeras columnas
                                for col in df.columns[:5]:
                                    valor = str(fila[col]).strip()
                                    if len(valor) <= 4 and valor.isalpha():
                                        inicial = valor
                                        break
                                
                                if inicial:
                                    iniciales_con_trop.append({
                                        'inicial': inicial,
                                        'columna_sat': col_sat,
                                        'fila': idx + 1,
                                        'valor_completo': str(fila[col_sat])
                                    })
                                    logging.info(f"Encontrado TROP: {inicial} en {col_sat}")
            
            return iniciales_con_trop, numero_semana
            
        except Exception as e:
            logging.error(f"Error al procesar {archivo_excel}: {e}")
            return [], None

    def actualizar_historial_csv(self, iniciales_con_trop, numero_semana):
        """Actualiza el archivo CSV con la semana para las personas que tuvieron TROP."""
        try:
            # Crear archivo CSV si no existe
            if not os.path.exists(self.archivo_csv):
                df_historial = pd.DataFrame({'empleado': [], 'ultima_semana_trop_sabado': []})
                df_historial.to_csv(self.archivo_csv, index=False)
                logging.info(f"Archivo CSV creado: {self.archivo_csv}")
            
            # Leer el archivo CSV actual
            df_historial = pd.read_csv(self.archivo_csv)
            logging.info(f"Archivo CSV actual cargado con {len(df_historial)} registros")
            
            # Obtener iniciales encontradas
            iniciales_encontradas = [item['inicial'] for item in iniciales_con_trop]
            logging.info(f"Iniciales con TROP: {iniciales_encontradas}")
            
            # Actualizar registros existentes y agregar nuevos
            actualizaciones = 0
            nuevas_entradas = 0
            
            for inicial in iniciales_encontradas:
                # Buscar si ya existe el empleado
                empleado_existente = df_historial[df_historial['empleado'] == inicial]
                
                if not empleado_existente.index.empty:
                    # Actualizar semana existente
                    idx = empleado_existente.index[0]
                    df_historial.at[idx, 'ultima_semana_trop_sabado'] = numero_semana
                    actualizaciones += 1
                    logging.info(f"Actualizado: {inicial} -> Semana {numero_semana}")
                else:
                    # Agregar nuevo empleado
                    nueva_fila = {'empleado': inicial, 'ultima_semana_trop_sabado': numero_semana}
                    df_historial = pd.concat([df_historial, pd.DataFrame([nueva_fila])], ignore_index=True)
                    nuevas_entradas += 1
                    logging.info(f"Nuevo empleado agregado: {inicial} -> Semana {numero_semana}")
            
            # Convertir la columna a enteros (manteniendo NaN para valores vacíos)
            df_historial['ultima_semana_trop_sabado'] = pd.to_numeric(df_historial['ultima_semana_trop_sabado'], errors='coerce')
            
            # Guardar archivo actualizado con números enteros
            df_historial.to_csv(self.archivo_csv, index=False, float_format='%.0f')
            
            logging.info(f"Archivo CSV actualizado: {actualizaciones} actualizaciones, {nuevas_entradas} nuevas entradas")
            
            return df_historial
            
        except Exception as e:
            logging.error(f"Error al actualizar CSV: {e}")
            return None

    def procesar_archivo_excel(self, archivo_excel):
        """Procesa un archivo Excel y actualiza el CSV."""
        try:
            logging.info(f"Iniciando procesamiento de: {archivo_excel}")
            
            # Extraer datos del Excel
            iniciales_con_trop, numero_semana = self.extraer_iniciales_con_trop(archivo_excel)
            
            if iniciales_con_trop and numero_semana:
                # Actualizar CSV
                df_actualizado = self.actualizar_historial_csv(iniciales_con_trop, numero_semana)
                
                if df_actualizado is not None:
                    logging.info("✅ Procesamiento completado exitosamente")
                    logging.info(f" Semana {numero_semana} asignada a {len(iniciales_con_trop)} empleados")
                    
                    # Mostrar resumen
                    print(f"\n{'='*60}")
                    print(f"✅ PROCESAMIENTO COMPLETADO")
                    print(f"{'='*60}")
                    print(f" Archivo procesado: {os.path.basename(archivo_excel)}")
                    print(f"📅 Semana: {numero_semana}")
                    print(f" Empleados con TROP: {len(iniciales_con_trop)}")
                    print(f"📝 Iniciales: {', '.join([item['inicial'] for item in iniciales_con_trop])}")
                    print(f" CSV actualizado: {self.archivo_csv}")
                    print(f"{'='*60}\n")
                    
                    return True
            else:
                logging.warning("No se encontraron datos válidos para procesar")
                return False
                
        except Exception as e:
            logging.error(f"Error en procesamiento: {e}")
            return False

    def monitorear_carpeta(self):
        """Monitorea la carpeta en busca de nuevos archivos Excel."""
        logging.info(" Iniciando monitoreo de carpeta...")
        print(" MONITOR DE TROP EN SÁBADOS INICIADO")
        print(f"📁 Monitoreando: {self.carpeta_monitoreo}")
        print(f"💾 Actualizando: {self.archivo_csv}")
        print("⏳ Esperando archivos Excel... (Ctrl+C para salir)")
        print("="*60)
        
        try:
            while True:
                # Buscar archivos Excel en la carpeta
                archivos_excel = []
                for archivo in os.listdir(self.carpeta_monitoreo):
                    if archivo.endswith(('.xlsx', '.xls')) and archivo not in self.archivos_procesados:
                        archivos_excel.append(archivo)
                
                # Procesar archivos nuevos
                for archivo in archivos_excel:
                    ruta_completa = os.path.join(self.carpeta_monitoreo, archivo)
                    
                    logging.info(f"Nuevo archivo detectado: {archivo}")
                    print(f"\n📥 Nuevo archivo detectado: {archivo}")
                    
                    # Procesar archivo
                    if self.procesar_archivo_excel(ruta_completa):
                        # Marcar como procesado
                        self.archivos_procesados.add(archivo)
                        logging.info(f"Archivo procesado exitosamente: {archivo}")
                    else:
                        logging.error(f"Error al procesar archivo: {archivo}")
                
                # Esperar antes de la siguiente verificación
                time.sleep(5)  # Verificar cada 5 segundos
                
        except KeyboardInterrupt:
            logging.info("Monitoreo detenido por el usuario")
            print("\n Monitoreo detenido")

def main():
    monitor = TropMonitor()
    monitor.monitorear_carpeta()

if __name__ == "__main__":
    main() 