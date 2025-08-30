# Clase ExcelConverter - Documentación Completa

## 📋 Descripción

La clase `ExcelConverter` es una herramienta completa y robusta para convertir archivos Excel a DataFrames de pandas y viceversa. Proporciona una interfaz orientada a objetos que facilita el manejo de datos Excel con funcionalidades avanzadas de validación, limpieza y análisis.

## 🚀 Características Principales

### ✅ **Conversión Bidireccional**
- **Excel → DataFrame**: Carga archivos Excel con validación automática
- **DataFrame → Excel**: Exporta DataFrames a archivos Excel con opciones configurables

### ✅ **Validación Robusta**
- Verificación de existencia de archivos
- Validación de extensiones de archivo (.xlsx, .xls, .xlsm, .xlsb)
- Manejo de permisos y rutas inválidas

### ✅ **Limpieza de Datos**
- Eliminación automática de filas duplicadas
- Eliminación de columnas completamente vacías
- Manejo de valores nulos

### ✅ **Análisis y Estadísticas**
- Información detallada de DataFrames
- Estadísticas de memoria y tipos de datos
- Detección de valores nulos

### ✅ **Manejo de Errores**
- Excepciones específicas para diferentes tipos de errores
- Logging configurable
- Mensajes informativos con emojis

## 📦 Instalación

```bash
pip install pandas openpyxl xlrd
```

## 🎯 Uso Básico

### Importar la clase
```python
from excel_converter import ExcelConverter
```

### Crear instancia
```python
# Con mensajes informativos
converter = ExcelConverter(verbose=True)

# Sin mensajes informativos
converter = ExcelConverter(verbose=False)
```

### Cargar Excel a DataFrame
```python
# Carga básica
df = converter.cargar_excel("archivo.xlsx")

# Carga con limpieza automática
df = converter.convertir_excel_a_dataframe("archivo.xlsx", limpiar=True)

# Carga con parámetros adicionales
df = converter.cargar_excel("archivo.xlsx", sheet_name="Hoja1", skiprows=2)
```

### Exportar DataFrame a Excel
```python
# Exportación básica
converter.exportar_excel(df, "salida.xlsx")

# Exportación con opciones
converter.convertir_dataframe_a_excel(df, "salida.xlsx", 
                                     sheet_name="Datos", 
                                     index=False)
```

## 🔧 Métodos Principales

### `__init__(verbose=True)`
Inicializa el convertidor.

**Parámetros:**
- `verbose` (bool): Si mostrar mensajes informativos

### `validar_ruta_archivo(ruta, debe_existir=True)`
Valida una ruta de archivo.

**Parámetros:**
- `ruta` (str): Ruta a validar
- `debe_existir` (bool): Si verificar existencia del archivo

**Retorna:** `bool`

### `cargar_excel(ruta_archivo, **kwargs)`
Carga un archivo Excel a DataFrame.

**Parámetros:**
- `ruta_archivo` (str): Ruta del archivo Excel
- `**kwargs`: Argumentos adicionales para `pd.read_excel()`

**Retorna:** `pd.DataFrame`

### `exportar_excel(df, ruta_salida, sheet_name='Sheet1', index=False, **kwargs)`
Exporta un DataFrame a archivo Excel.

**Parámetros:**
- `df` (pd.DataFrame): DataFrame a exportar
- `ruta_salida` (str): Ruta de salida
- `sheet_name` (str): Nombre de la hoja
- `index` (bool): Si incluir índice
- `**kwargs`: Argumentos adicionales para `df.to_excel()`

**Retorna:** `bool`

### `mostrar_informacion(df=None)`
Muestra información detallada del DataFrame.

**Parámetros:**
- `df` (pd.DataFrame, optional): DataFrame a analizar

### `obtener_estadisticas(df=None)`
Obtiene estadísticas del DataFrame.

**Parámetros:**
- `df` (pd.DataFrame, optional): DataFrame a analizar

**Retorna:** `Dict[str, Any]`

### `limpiar_dataframe(df=None, eliminar_duplicados=True, eliminar_columnas_vacias=True)`
Limpia el DataFrame.

**Parámetros:**
- `df` (pd.DataFrame, optional): DataFrame a limpiar
- `eliminar_duplicados` (bool): Si eliminar filas duplicadas
- `eliminar_columnas_vacias` (bool): Si eliminar columnas vacías

**Retorna:** `pd.DataFrame`

## 📝 Ejemplos de Uso

### Ejemplo 1: Conversión Básica
```python
from excel_converter import ExcelConverter

# Crear convertidor
converter = ExcelConverter()

# Cargar archivo Excel
df = converter.convertir_excel_a_dataframe("datos.xlsx", limpiar=True)

# Mostrar información
converter.mostrar_informacion(df)

# Exportar a nuevo archivo
converter.convertir_dataframe_a_excel(df, "datos_procesados.xlsx")
```

### Ejemplo 2: Procesamiento Avanzado
```python
from excel_converter import ExcelConverter
import pandas as pd

converter = ExcelConverter(verbose=True)

# Cargar múltiples hojas
df1 = converter.cargar_excel("archivo.xlsx", sheet_name="Hoja1")
df2 = converter.cargar_excel("archivo.xlsx", sheet_name="Hoja2")

# Limpiar datos
df1_limpio = converter.limpiar_dataframe(df1)
df2_limpio = converter.limpiar_dataframe(df2)

# Combinar DataFrames
df_combinado = pd.concat([df1_limpio, df2_limpio], ignore_index=True)

# Exportar resultado
converter.exportar_excel(df_combinado, "resultado_final.xlsx")
```

### Ejemplo 3: Análisis de Datos
```python
from excel_converter import ExcelConverter

converter = ExcelConverter()

# Cargar datos
df = converter.cargar_excel("ventas.xlsx")

# Obtener estadísticas
stats = converter.obtener_estadisticas(df)

print(f"Dimensiones: {stats['dimensiones']}")
print(f"Memoria utilizada: {stats['memoria_mb']:.2f} MB")
print(f"Columnas numéricas: {stats['columnas_numericas']}")
print(f"Valores nulos: {stats['valores_nulos']}")
```

### Ejemplo 4: Manejo de Errores
```python
from excel_converter import ExcelConverter

converter = ExcelConverter()

try:
    # Intentar cargar archivo
    df = converter.cargar_excel("archivo_inexistente.xlsx")
except FileNotFoundError:
    print("Archivo no encontrado")
except PermissionError:
    print("Sin permisos para acceder al archivo")
except Exception as e:
    print(f"Error inesperado: {e}")
```

## 🛡️ Manejo de Errores

La clase maneja los siguientes tipos de errores:

- **FileNotFoundError**: Archivo no encontrado
- **PermissionError**: Permisos insuficientes
- **ValueError**: Ruta inválida o formato incorrecto
- **Exception**: Otros errores de lectura/escritura

## 📊 Atributos de la Clase

- `verbose` (bool): Modo verboso
- `extensiones_validas` (list): Extensiones de archivo válidas
- `ultimo_dataframe` (pd.DataFrame): Último DataFrame cargado
- `ultima_ruta` (str): Última ruta procesada
- `logger` (logging.Logger): Logger configurado

## 🔍 Validación de Archivos

La clase valida automáticamente:

- ✅ Existencia del archivo
- ✅ Extensión válida (.xlsx, .xls, .xlsm, .xlsb)
- ✅ Que sea un archivo (no directorio)
- ✅ Permisos de acceso

## 🧹 Limpieza Automática

Opciones de limpieza disponibles:

- **Eliminar duplicados**: Remueve filas duplicadas
- **Eliminar columnas vacías**: Remueve columnas completamente vacías
- **Manejo de valores nulos**: Detecta y reporta valores nulos

## 📈 Información Mostrada

La función `mostrar_informacion()` muestra:

- 📏 Dimensiones del DataFrame
- 📋 Lista de columnas
- 🔍 Tipos de datos
- ⚠️ Valores nulos por columna
- 👀 Primeras y últimas filas
- 💾 Uso de memoria

## 🚀 Ventajas de Usar la Clase

### ✅ **Reutilización**
- Una instancia puede procesar múltiples archivos
- Mantiene estado del último DataFrame procesado

### ✅ **Flexibilidad**
- Parámetros configurables para cada operación
- Soporte para argumentos adicionales de pandas

### ✅ **Robustez**
- Manejo completo de errores
- Validación automática de archivos

### ✅ **Facilidad de Uso**
- Métodos de conveniencia para operaciones comunes
- Interfaz intuitiva con mensajes informativos

### ✅ **Extensibilidad**
- Fácil de extender con nuevas funcionalidades
- Compatible con el ecosistema de pandas

## 📁 Archivos del Proyecto

- `excel_converter.py`: Clase principal
- `ejemplo_uso_clase.py`: Ejemplos de uso
- `excel_to_dataframe.py`: Script original (funcional)
- `crear_archivo_ejemplo.py`: Generador de datos de prueba
- `requirements.txt`: Dependencias
- `README.md`: Documentación general
- `README_CLASE.md`: Esta documentación

## 🤝 Contribuciones

Las contribuciones son bienvenidas. Algunas ideas para mejorar:

- Soporte para más formatos de archivo
- Funciones de transformación de datos
- Integración con bases de datos
- Interfaz gráfica
- Procesamiento en lotes

## 📄 Licencia

Este proyecto es de código abierto y está disponible bajo la licencia MIT. 