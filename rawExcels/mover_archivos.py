import time
import os
import shutil
import subprocess
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import xlwings as xw

ORIGEN = r"C:\Users\Usuario1\Desktop\cursor\excel_extract\excel_extraction_forschedule"
DESTINO = r"C:\Users\Usuario1\Desktop\horario\rawExcels"

class MoverArchivosHandler(FileSystemEventHandler):
    def on_created(self, event):
        if not event.is_directory:
            filename = os.path.basename(event.src_path)
            if filename.startswith("horario_descansos_semana") and filename.endswith(".xlsx"):
                try:
                    destino_path = os.path.join(DESTINO, filename)
                    # Espera a que el archivo termine de copiarse/escribirse
                    time.sleep(2)
                    if os.path.exists(event.src_path):
                        shutil.move(event.src_path, destino_path)
                        print(f"✅ Movido: {filename}")
                        
                        # Copiar datos al archivo horioUnificado.xlsm
                        try:
                            # Extraer el nombre de la hoja (parte después de "horario_descansos_semana_")
                            nombre_base = os.path.splitext(filename)[0]
                            nombre_hoja = nombre_base.replace("horario_descansos_semana_", "")
                            
                            # Ruta del archivo unificado
                            archivo_unificado = os.path.join(DESTINO, "horioUnificado.xlsm")
                            
                            if os.path.exists(archivo_unificado):
                                # Abrir el archivo unificado con xlwings
                                app = xw.App(visible=False)
                                wb_unificado = app.books.open(archivo_unificado)
                                
                                # Verificar si la hoja ya existe, si no, crearla
                                if nombre_hoja not in [sheet.name for sheet in wb_unificado.sheets]:
                                    wb_unificado.sheets.add(name=nombre_hoja)
                                
                                # Seleccionar la hoja
                                hoja = wb_unificado.sheets[nombre_hoja]
                                
                                # Limpiar la hoja existente
                                hoja.used_range.clear_contents()
                                
                                # Copiar datos del archivo original
                                wb_origen = app.books.open(destino_path)
                                hoja_origen = wb_origen.sheets[0]  # Primera hoja
                                
                                # Copiar todo el contenido usado
                                rango_origen = hoja_origen.used_range
                                rango_origen.copy(hoja.range('A1'))
                                
                                # Guardar el archivo unificado
                                wb_unificado.save()
                                wb_origen.close()
                                wb_unificado.close()
                                app.quit()
                                
                                print(f"📋 Copiado a hoja '{nombre_hoja}' en horioUnificado.xlsm")
                                
                                # Abrir el archivo unificado
                                subprocess.run(['start', '', archivo_unificado], shell=True)
                                print(f"📂 Abriendo horioUnificado.xlsm")
                                
                            else:
                                print(f"⚠️  El archivo horioUnificado.xlsm no existe en {DESTINO}")
                                # Abrir el archivo original
                                subprocess.run(['start', '', destino_path], shell=True)
                                print(f"📂 Abriendo archivo original: {filename}")
                                
                        except Exception as e:
                            print(f"❌ Error copiando a horioUnificado.xlsm: {e}")
                            # Si falla, abrir el archivo original
                            subprocess.run(['start', '', destino_path], shell=True)
                            print(f"📂 Abriendo archivo original: {filename}")
                    else:
                        print(f"❌ Error: El archivo {filename} ya no existe en origen")
                except Exception as e:
                    print(f"❌ Error moviendo {filename}: {e}")

if __name__ == "__main__":
    print(f"📁 Monitoreando carpeta: {ORIGEN}")
    print(f"📁 Destino: {DESTINO}")
    print("⏳ Esperando nuevos archivos... (Ctrl+C para detener)")
    
    event_handler = MoverArchivosHandler()
    observer = Observer()
    observer.schedule(event_handler, ORIGEN, recursive=False)
    observer.start()
    
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("\n🛑 Deteniendo monitoreo...")
        observer.stop()
    observer.join() 