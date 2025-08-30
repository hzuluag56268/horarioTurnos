# 📋 Sistema de Restricciones Externas

## 🎯 Descripción General

El **Sistema de Restricciones Externas** permite configurar todas las restricciones, turnos especiales, fechas específicas y configuraciones del generador de turnos desde un archivo externo (`config_restricciones.py`), sin necesidad de modificar el código principal.

## 🏗️ Arquitectura del Sistema

```
📁 Proyecto/
├── 📄 config_restricciones.py          # ⭐ Configuración Externa
├── 📄 generador_descansos_separacion.py # Generador Principal
├── 📄 test_restricciones_externas.py    # Script de Pruebas
└── 📄 README_RESTRICCIONES_EXTERNAS.md  # Esta documentación
```

## 🔧 Componentes Principales

### 1. **Archivo de Configuración Externa** (`config_restricciones.py`)

#### 📋 **Restricciones de Empleados (DESC/TROP)**
```python
RESTRICCIONES_EMPLEADOS = {
    "HZG": {
        "DESC": {"dias_permitidos": ["martes"], "tipo": "fijo"},
        "TROP": {"dias_permitidos": ["miércoles", "jueves", "viernes", "sábado"], "tipo": "opcional"}
    },
    "ROP": {
        "DESC": {"libre": True},  # Sin restricción
        "TROP": {"libre": True}   # Sin restricción
    }
}
```

#### 📅 **Turnos por Fechas Específicas (Máxima Prioridad)**
```python
TURNOS_FECHAS_ESPECIFICAS = {
    "JIS": [
        {"fecha": "2025-07-17", "turno_requerido": "VACA"},
        {"fecha": "2025-07-18", "turno_requerido": "VACA"},
        # ... más fechas
    ],
    "AFG": [
        {"fecha": "2025-07-01", "turno_requerido": "COME"},
        # ... más fechas
    ]
}
```

#### 🎯 **Turnos Especiales Extendidos**
```python
TURNOS_ESPECIALES = {
    "GMT": [
        {
            "tipo": "SIND",
            "frecuencia": "semanal_fijo",
            "dia_semana": "miércoles"
        }
    ]
}
```

#### 🚫 **Trabajadores Fuera de Operación**
```python
TRABAJADORES_FUERA_OPERACION = ['PHD', 'WEH', 'VCM', 'MEI']
```

#### 🎉 **Días Festivos (2025-2030)**
```python
DIAS_FESTIVOS = [
    "2025-01-01",  # Año Nuevo
    "2025-07-20",  # Día de la Independencia
    # ... más fechas hasta 2030
]
```

#### ⚙️ **Configuración General**
```python
CONFIGURACION_GENERAL = {
    "archivo_historial_sabados": "historial_sabados.csv",
    "mapeo_dias": {
        "lunes": 0, "martes": 1, "miércoles": 2, "jueves": 3, 
        "viernes": 4, "sábado": 5, "domingo": 6
    },
    "turnos_validos": ["DESC", "TROP", "SIND", "VACA", "COME", "COMT", "COMS", "CMED", "CERT"],
    "turnos_completos": ["VACA", "COME", "COMT", "COMS"],  # Reemplazan DESC/TROP
    "turnos_adicionales": ["SIND", "CMED", "CERT"]         # Se suman a DESC/TROP
}
```

## 🔍 Funciones de Validación y Utilidad

### **Validación Automática**
```python
# Validar configuración
errores = validar_configuracion()
if errores:
    print("❌ Errores encontrados:", errores)

# Obtener resumen
resumen = obtener_resumen_configuracion()
```

### **Funciones de Modificación Dinámica**
```python
# Agregar nueva restricción
agregar_restriccion_empleado("NUEVO", "DESC", ["lunes"], "fijo")

# Agregar fecha específica
agregar_fecha_especifica("NUEVO", "2025-08-01", "VACA")

# Agregar turno especial
agregar_turno_especial("NUEVO", "CERT", "semanal_fijo", "viernes")
```

## 🚀 Uso del Sistema

### 1. **Uso Básico**
```python
from generador_descansos_separacion import GeneradorDescansosSeparacion

# Crear generador (automáticamente carga restricciones externas)
generador = GeneradorDescansosSeparacion(
    año=2025,
    mes=7,
    num_empleados=24,
    semana_especifica=28
)

# Generar horario
df = generador.generar_horario_primera_semana()
```

### 2. **Modificar Configuración**
```python
# Importar configuración
from config_restricciones import (
    RESTRICCIONES_EMPLEADOS,
    agregar_restriccion_empleado
)

# Agregar nueva restricción
agregar_restriccion_empleado("NUEVO_EMP", "DESC", ["miércoles"], "fijo")

# Usar generador con nueva configuración
generador = GeneradorDescansosSeparacion(año=2025, semana_especifica=30)
```

## 🧪 Sistema de Pruebas

### **Ejecutar Validación de Configuración**
```bash
python config_restricciones.py
```

### **Ejecutar Pruebas Completas**
```bash
python test_restricciones_externas.py
```

### **Pruebas Incluidas**
1. ✅ **Importación de Configuración Externa**
2. ✅ **Generador con Restricciones Externas**
3. ✅ **Generación de Horario Completo**
4. ✅ **Modificación Dinámica de Configuración**

## 📊 Tipos de Turnos Soportados

### **Turnos Básicos**
- **DESC**: Descanso
- **TROP**: Turno Operacional

### **Turnos Especiales Completos** (Reemplazan DESC/TROP)
- **VACA**: Vacaciones
- **COME**: Comisión Externa
- **COMT**: Comisión de Trabajo
- **COMS**: Comisión de Servicio

### **Turnos Adicionales** (Se suman a DESC/TROP)
- **SIND**: Sindicato
- **CMED**: Cita Médica
- **CERT**: Certificación

## 🎯 Características Principales

### ✅ **Ventajas del Sistema Externo**

1. **🔧 Fácil Mantenimiento**
   - Modificar restricciones sin tocar código principal
   - Separación clara de configuración y lógica

2. **🛡️ Validación Automática**
   - Verificación de integridad al cargar
   - Detección temprana de errores

3. **🔄 Modificación Dinámica**
   - Agregar/modificar restricciones en tiempo de ejecución
   - Funciones de utilidad incluidas

4. **📋 Configuración Centralizada**
   - Todas las restricciones en un solo lugar
   - Fácil backup y versionado

5. **🧪 Sistema de Pruebas Integrado**
   - Validación automática completa
   - Pruebas de integración

### 🎨 **Flexibilidad de Configuración**

- **Restricciones por Empleado**: Días específicos para DESC/TROP
- **Fechas Exactas**: Turnos obligatorios en fechas específicas
- **Turnos Especiales**: Configuración avanzada (SIND, CERT, etc.)
- **Exclusiones**: Trabajadores fuera de operación
- **Días Festivos**: Configuración hasta 2030

## 📈 Ejemplos de Configuración

### **Ejemplo 1: Empleado con Restricción Fija**
```python
"HZG": {
    "DESC": {"dias_permitidos": ["martes"], "tipo": "fijo"},
    "TROP": {"dias_permitidos": ["miércoles", "jueves", "viernes", "sábado"], "tipo": "opcional"}
}
```

### **Ejemplo 2: Vacaciones Extendidas**
```python
"JIS": [
    {"fecha": "2025-07-17", "turno_requerido": "VACA"},
    {"fecha": "2025-07-18", "turno_requerido": "VACA"},
    {"fecha": "2025-07-19", "turno_requerido": "VACA"},
    # ... 18 días de vacaciones
]
```

### **Ejemplo 3: Turno Especial Semanal**
```python
"GMT": [
    {
        "tipo": "SIND",
        "frecuencia": "semanal_fijo",
        "dia_semana": "miércoles"
    }
]
```

## 🔧 Mantenimiento y Actualizaciones

### **Agregar Nuevo Empleado**
1. Agregar a `RESTRICCIONES_EMPLEADOS` si tiene restricciones
2. Agregar a `TURNOS_FECHAS_ESPECIFICAS` si tiene fechas específicas
3. Agregar a `TURNOS_ESPECIALES` si tiene turnos especiales
4. Ejecutar `python config_restricciones.py` para validar

### **Agregar Nuevo Tipo de Turno**
1. Agregar a `CONFIGURACION_GENERAL["turnos_validos"]`
2. Clasificar en `turnos_completos` o `turnos_adicionales`
3. Actualizar lógica en generador principal si es necesario

### **Actualizar Días Festivos**
1. Agregar fechas a `DIAS_FESTIVOS`
2. Mantener formato "YYYY-MM-DD"
3. Validar con `python config_restricciones.py`

## 🎉 Resultado Final

### **✅ Sistema Completamente Funcional**
- ✅ Configuración externa implementada
- ✅ Validación automática funcionando
- ✅ Generador principal actualizado
- ✅ Todas las pruebas pasando
- ✅ Documentación completa

### **🎯 Beneficios Obtenidos**
1. **Mantenimiento Simplificado**: Cambios sin modificar código principal
2. **Configuración Centralizada**: Todo en un solo archivo
3. **Validación Robusta**: Detección automática de errores
4. **Flexibilidad Total**: Modificaciones dinámicas soportadas
5. **Pruebas Integradas**: Validación completa automatizada

---

## 📞 Soporte

Para modificar restricciones, editar el archivo `config_restricciones.py` y ejecutar:
```bash
python config_restricciones.py  # Validar configuración
python test_restricciones_externas.py  # Probar sistema completo
```

¡El sistema está listo para usar! 🚀 