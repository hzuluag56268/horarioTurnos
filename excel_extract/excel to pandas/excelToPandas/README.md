# Conversor de Excel a DataFrame

Este script de Python convierte archivos Excel a DataFrames de pandas de manera interactiva y robusta.

## 🚀 Características

- ✅ **Validación de rutas**: Verifica que el archivo exista y tenga formato Excel válido
- ✅ **Manejo de errores robusto**: Captura y maneja diferentes tipos de errores
- ✅ **Información detallada**: Muestra dimensiones, columnas, tipos de datos y más
- ✅ **Interfaz amigable**: Interfaz de consola con emojis y mensajes claros
- ✅ **Soporte múltiples formatos**: Compatible con .xlsx, .xls, .xlsm, .xlsb
- ✅ **Código limpio y comentado**: Fácil de entender y mantener

## 📋 Requisitos

- Python 3.7 o superior
- pandas
- openpyxl (para archivos .xlsx)
- xlrd (para archivos .xls)

## 🔧 Instalación

1. **Clona o descarga este proyecto**

2. **Instala las dependencias**:
   ```bash
   pip install -r requirements.txt
   ```

   O instala manualmente:
   ```bash
   pip install pandas openpyxl xlrd
   ```

## 🎯 Uso

1. **Ejecuta el script**:
   ```bash
   python excel_to_dataframe.py
   ```

2. **Ingresa la ruta del archivo Excel** cuando se te solicite:
   ```
   ➤ Ruta del archivo: C:\Users\Usuario\Documentos\datos.xlsx
   ```

3. **Revisa la información del DataFrame** que se muestra automáticamente

4. **Decide si quieres convertir otro archivo** cuando se te pregunte

## 📊 Información mostrada

El script muestra la siguiente información del DataFrame:

- **Dimensiones**: Número de filas y columnas
- **Columnas**: Lista de todas las columnas
- **Tipos de datos**: Tipo de cada columna
- **Valores nulos**: Cantidad de valores nulos por columna
- **Vista previa**: Primeras y últimas 5 filas
- **Uso de memoria**: Memoria utilizada por el DataFrame

## 🛡️ Manejo de errores

El script maneja los siguientes tipos de errores:

- ❌ Archivo no encontrado
- ❌ Permisos insuficientes
- ❌ Formato de archivo no válido
- ❌ Rutas vacías o inválidas
- ❌ Errores de lectura del archivo Excel

## 📝 Ejemplo de uso

```
🚀 CONVERSOR DE EXCEL A DATAFRAME
==================================================
Este programa convierte archivos Excel a DataFrames de pandas
==================================================

📁 Por favor, ingresa la ruta completa del archivo Excel:
   Ejemplo: C:\Users\Usuario\Documentos\archivo.xlsx
   O: /home/usuario/documentos/archivo.xlsx

➤ Ruta del archivo: datos.xlsx

📂 Cargando archivo: datos.xlsx
✅ Archivo Excel cargado exitosamente!

============================================================
📊 INFORMACIÓN DEL DATAFRAME
============================================================
📏 Dimensiones: 1000 filas × 5 columnas

📋 Columnas (5):
   1. Nombre
   2. Edad
   3. Ciudad
   4. Salario
   5. Fecha

🔍 Tipos de datos:
   Nombre: object
   Edad: int64
   Ciudad: object
   Salario: float64
   Fecha: datetime64[ns]

✅ No hay valores nulos en el DataFrame

👀 Primeras 5 filas del DataFrame:
----------------------------------------
     Nombre  Edad    Ciudad    Salario      Fecha
0    Juan    25    Madrid    45000.0  2023-01-15
1    María   30   Barcelona  52000.0  2023-01-16
2    Pedro   28    Valencia  48000.0  2023-01-17
3    Ana     35     Sevilla  55000.0  2023-01-18
4    Carlos  27     Málaga   47000.0  2023-01-19

💾 Uso de memoria: 0.15 MB

============================================================
🎉 ¡CONVERSIÓN EXITOSA!
============================================================
✅ El archivo 'datos.xlsx' ha sido convertido
✅ DataFrame creado con 1000 filas y 5 columnas
============================================================

¿Deseas convertir otro archivo? (s/n):
➤ n

👋 ¡Gracias por usar el conversor! Hasta luego.
```

## 🔍 Funciones principales

### `validar_ruta_archivo(ruta)`
Valida que la ruta ingresada sea válida y el archivo exista.

### `cargar_excel_a_dataframe(ruta_archivo)`
Carga un archivo Excel y lo convierte a un DataFrame de pandas.

### `mostrar_informacion_dataframe(df)`
Muestra información detallada del DataFrame cargado.

### `main()`
Función principal que coordina todo el proceso.

## 🐛 Solución de problemas

### Error: "No module named 'pandas'"
```bash
pip install pandas
```

### Error: "No module named 'openpyxl'"
```bash
pip install openpyxl
```

### Error: "No module named 'xlrd'"
```bash
pip install xlrd
```

### Error al leer archivos .xls
Asegúrate de tener instalado `xlrd`:
```bash
pip install xlrd
```

## 📄 Licencia

Este proyecto es de código abierto y está disponible bajo la licencia MIT.

## 🤝 Contribuciones

Las contribuciones son bienvenidas. Por favor, abre un issue o pull request para sugerir mejoras. 