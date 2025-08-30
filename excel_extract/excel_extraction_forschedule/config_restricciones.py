#!/usr/bin/env python3
"""
Configuración de Restricciones - Sistema de Turnos
==================================================
Archivo simple para configurar restricciones del generador de turnos.
"""

# ============================================================================
# RESTRICCIONES DE EMPLEADOS
# ============================================================================

RESTRICCIONES_EMPLEADOS = {
    #Dejare estas líneas como ejemplo.
    #"ROP": {
    #    "DESC": {"dias_permitidos": ["martes"], "tipo": "fijo"},
    #    "TROP": {"dias_permitidos": ["miércoles", "jueves", "viernes", "sábado"], "tipo": "opcional"}
    #},
    "HZG": {
        "DESC": {"libre": True},
        "TROP": {"libre": True}
    }
}

# ============================================================================
# FECHAS ESPECÍFICAS
# ============================================================================

TURNOS_FECHAS_ESPECIFICAS = {
    "JIS": [
        {"fecha": "2025-07-14", "turno_requerido": "VACA"},
        {"fecha": "2025-07-15", "turno_requerido": "VACA"},
        {"fecha": "2025-07-16", "turno_requerido": "VACA"},
        {"fecha": "2025-07-17", "turno_requerido": "VACA"},
        {"fecha": "2025-07-18", "turno_requerido": "VACA"},
        {"fecha": "2025-07-19", "turno_requerido": "VACA"},
        {"fecha": "2025-07-20", "turno_requerido": "VACA"},
        {"fecha": "2025-07-21", "turno_requerido": "VACA"},
        {"fecha": "2025-07-22", "turno_requerido": "VACA"},
        {"fecha": "2025-07-23", "turno_requerido": "VACA"},
        {"fecha": "2025-07-24", "turno_requerido": "VACA"},
        {"fecha": "2025-07-25", "turno_requerido": "VACA"},
        {"fecha": "2025-07-26", "turno_requerido": "VACA"},
        {"fecha": "2025-07-27", "turno_requerido": "VACA"},
        {"fecha": "2025-07-28", "turno_requerido": "VACA"},
        {"fecha": "2025-07-29", "turno_requerido": "VACA"},
        {"fecha": "2025-07-30", "turno_requerido": "VACA"},
        {"fecha": "2025-07-31", "turno_requerido": "VACA"},
        {"fecha": "2025-08-01", "turno_requerido": "VACA"},
    ],
    "AFG": [
        {"fecha": "2025-07-01", "turno_requerido": "COME"},
        {"fecha": "2025-07-02", "turno_requerido": "COME"},
        {"fecha": "2025-07-03", "turno_requerido": "COME"},
        {"fecha": "2025-07-04", "turno_requerido": "COME"},
        {"fecha": "2025-07-05", "turno_requerido": "COME"},
        {"fecha": "2025-07-06", "turno_requerido": "COME"},
        {"fecha": "2025-07-07", "turno_requerido": "COME"},
        {"fecha": "2025-07-08", "turno_requerido": "COME"},
        {"fecha": "2025-07-09", "turno_requerido": "COME"},
        {"fecha": "2025-07-10", "turno_requerido": "COME"},
        {"fecha": "2025-07-11", "turno_requerido": "COME"},
        {"fecha": "2025-07-12", "turno_requerido": "COME"},
        {"fecha": "2025-07-13", "turno_requerido": "COME"},
        {"fecha": "2025-07-14", "turno_requerido": "COME"},
        {"fecha": "2025-07-15", "turno_requerido": "COME"},
        {"fecha": "2025-07-16", "turno_requerido": "COME"},
        {"fecha": "2025-07-17", "turno_requerido": "COME"},
        {"fecha": "2025-07-18", "turno_requerido": "COME"},
        {"fecha": "2025-07-19", "turno_requerido": "COME"},
        {"fecha": "2025-07-20", "turno_requerido": "COME"},
        {"fecha": "2025-07-21", "turno_requerido": "COME"},
        {"fecha": "2025-07-22", "turno_requerido": "COME"},
        {"fecha": "2025-07-23", "turno_requerido": "COME"},
        {"fecha": "2025-07-24", "turno_requerido": "COME"},
        {"fecha": "2025-07-25", "turno_requerido": "COME"},
        {"fecha": "2025-07-26", "turno_requerido": "COME"},
        {"fecha": "2025-07-27", "turno_requerido": "COME"},
        {"fecha": "2025-07-28", "turno_requerido": "COME"},
        {"fecha": "2025-07-29", "turno_requerido": "COME"},
        {"fecha": "2025-07-30", "turno_requerido": "COME"},
        {"fecha": "2025-07-31", "turno_requerido": "COME"},
        {"fecha": "2025-08-01", "turno_requerido": "COME"},
        {"fecha": "2025-08-02", "turno_requerido": "COME"},
        {"fecha": "2025-08-03", "turno_requerido": "COME"},
        {"fecha": "2025-08-04", "turno_requerido": "COME"},
        {"fecha": "2025-08-05", "turno_requerido": "COME"},
        {"fecha": "2025-08-06", "turno_requerido": "COME"},
        {"fecha": "2025-08-07", "turno_requerido": "COME"},
        {"fecha": "2025-08-08", "turno_requerido": "COME"},
        {"fecha": "2025-08-09", "turno_requerido": "COME"},
        {"fecha": "2025-08-10", "turno_requerido": "COME"},
        {"fecha": "2025-08-11", "turno_requerido": "COME"},
        {"fecha": "2025-08-12", "turno_requerido": "COME"},
        {"fecha": "2025-08-13", "turno_requerido": "COME"},
        {"fecha": "2025-08-14", "turno_requerido": "COME"},
        {"fecha": "2025-08-15", "turno_requerido": "COME"},
        {"fecha": "2025-08-16", "turno_requerido": "COME"},
        {"fecha": "2025-08-17", "turno_requerido": "COME"},
        {"fecha": "2025-08-18", "turno_requerido": "COME"},
        {"fecha": "2025-08-19", "turno_requerido": "COME"},
        {"fecha": "2025-08-20", "turno_requerido": "COME"},
        {"fecha": "2025-08-21", "turno_requerido": "COME"},
        {"fecha": "2025-08-22", "turno_requerido": "COME"},
        {"fecha": "2025-08-23", "turno_requerido": "COME"},
        {"fecha": "2025-08-24", "turno_requerido": "COME"},
        {"fecha": "2025-08-25", "turno_requerido": "COME"},
        {"fecha": "2025-08-26", "turno_requerido": "COME"},
        {"fecha": "2025-08-27", "turno_requerido": "COME"},
        {"fecha": "2025-08-28", "turno_requerido": "COME"},
        {"fecha": "2025-08-29", "turno_requerido": "COME"},
    ],
    "JMV": [
        {"fecha": "2025-07-22", "turno_requerido": "COMS"},
        {"fecha": "2025-07-23", "turno_requerido": "COMS"},
        {"fecha": "2025-07-24", "turno_requerido": "COMS"},
        {"fecha": "2025-07-25", "turno_requerido": "COMS"},
        {"fecha": "2025-07-26", "turno_requerido": "COMS"},
        {"fecha": "2025-07-27", "turno_requerido": "COMS"},
        {"fecha": "2025-07-28", "turno_requerido": "COMS"},
        {"fecha": "2025-07-29", "turno_requerido": "COMS"},
        {"fecha": "2025-07-30", "turno_requerido": "COMS"},
        {"fecha": "2025-07-31", "turno_requerido": "COMS"},
        {"fecha": "2025-08-01", "turno_requerido": "COMS"},
        {"fecha": "2025-08-02", "turno_requerido": "COMS"},
        {"fecha": "2025-08-03", "turno_requerido": "COMS"},
        {"fecha": "2025-09-22", "turno_requerido": "COME"},
        {"fecha": "2025-09-23", "turno_requerido": "COME"},
        {"fecha": "2025-09-24", "turno_requerido": "COME"},
        {"fecha": "2025-09-25", "turno_requerido": "COME"},
        {"fecha": "2025-09-26", "turno_requerido": "COME"},
        {"fecha": "2025-09-27", "turno_requerido": "COME"},
        {"fecha": "2025-09-28", "turno_requerido": "COME"},
        {"fecha": "2025-09-29", "turno_requerido": "COME"},
        {"fecha": "2025-09-30", "turno_requerido": "COME"},
        {"fecha": "2025-10-01", "turno_requerido": "COME"},
        {"fecha": "2025-10-02", "turno_requerido": "COME"},
        {"fecha": "2025-10-03", "turno_requerido": "COME"},
    ],
    "HLG": [
        {"fecha": "2025-07-18", "turno_requerido": "CMED"},
        {"fecha": "2025-09-16", "turno_requerido": "VACA"},
        {"fecha": "2025-09-17", "turno_requerido": "VACA"},
        {"fecha": "2025-09-18", "turno_requerido": "VACA"},
        {"fecha": "2025-09-19", "turno_requerido": "VACA"},
        {"fecha": "2025-09-20", "turno_requerido": "VACA"},
        {"fecha": "2025-09-21", "turno_requerido": "VACA"},
        {"fecha": "2025-09-22", "turno_requerido": "VACA"},
        {"fecha": "2025-09-23", "turno_requerido": "VACA"},
        {"fecha": "2025-09-24", "turno_requerido": "VACA"},
        {"fecha": "2025-09-25", "turno_requerido": "VACA"},
        {"fecha": "2025-09-26", "turno_requerido": "VACA"},
        {"fecha": "2025-09-27", "turno_requerido": "VACA"},
        {"fecha": "2025-09-28", "turno_requerido": "VACA"},
        {"fecha": "2025-09-29", "turno_requerido": "VACA"},
        {"fecha": "2025-09-30", "turno_requerido": "VACA"},
        {"fecha": "2025-10-01", "turno_requerido": "VACA"},
        {"fecha": "2025-10-02", "turno_requerido": "VACA"},
        {"fecha": "2025-10-03", "turno_requerido": "VACA"},
        {"fecha": "2025-10-04", "turno_requerido": "VACA"},
        {"fecha": "2025-10-05", "turno_requerido": "VACA"},
        {"fecha": "2025-10-06", "turno_requerido": "VACA"},
    ],
    "GMT": [
        {"fecha": "2025-09-22", "turno_requerido": "VACA"},
        {"fecha": "2025-09-23", "turno_requerido": "VACA"},
        {"fecha": "2025-09-24", "turno_requerido": "VACA"},
        {"fecha": "2025-09-25", "turno_requerido": "VACA"},
        {"fecha": "2025-09-26", "turno_requerido": "VACA"},
        {"fecha": "2025-09-27", "turno_requerido": "VACA"},
        {"fecha": "2025-09-28", "turno_requerido": "VACA"},
        {"fecha": "2025-09-29", "turno_requerido": "VACA"},
        {"fecha": "2025-09-30", "turno_requerido": "VACA"},
        {"fecha": "2025-10-01", "turno_requerido": "VACA"},
        {"fecha": "2025-10-02", "turno_requerido": "VACA"},
        {"fecha": "2025-10-03", "turno_requerido": "VACA"},
        {"fecha": "2025-10-04", "turno_requerido": "VACA"},
        {"fecha": "2025-10-05", "turno_requerido": "VACA"},
        {"fecha": "2025-10-06", "turno_requerido": "VACA"},
        {"fecha": "2025-10-07", "turno_requerido": "VACA"},
        {"fecha": "2025-10-08", "turno_requerido": "VACA"},
        {"fecha": "2025-10-09", "turno_requerido": "VACA"},
        {"fecha": "2025-10-10", "turno_requerido": "VACA"},
    ],
    "DJO": [
        {"fecha": "2025-09-22", "turno_requerido": "VACA"},
        {"fecha": "2025-09-23", "turno_requerido": "VACA"},
        {"fecha": "2025-09-24", "turno_requerido": "VACA"},
        {"fecha": "2025-09-25", "turno_requerido": "VACA"},
        {"fecha": "2025-09-26", "turno_requerido": "VACA"},
        {"fecha": "2025-09-27", "turno_requerido": "VACA"},
        {"fecha": "2025-09-28", "turno_requerido": "VACA"},
        {"fecha": "2025-09-29", "turno_requerido": "VACA"},
        {"fecha": "2025-09-30", "turno_requerido": "VACA"},
        {"fecha": "2025-10-01", "turno_requerido": "VACA"},
        {"fecha": "2025-10-02", "turno_requerido": "VACA"},
        {"fecha": "2025-10-03", "turno_requerido": "VACA"},
        {"fecha": "2025-10-04", "turno_requerido": "VACA"},
        {"fecha": "2025-10-05", "turno_requerido": "VACA"},
        {"fecha": "2025-10-06", "turno_requerido": "VACA"},
        {"fecha": "2025-10-07", "turno_requerido": "VACA"},
        {"fecha": "2025-10-08", "turno_requerido": "VACA"},
        {"fecha": "2025-10-09", "turno_requerido": "VACA"},
        {"fecha": "2025-10-10", "turno_requerido": "VACA"},
    ],
    "JLF": [
        {"fecha": "2025-09-01", "turno_requerido": "COME"},
        {"fecha": "2025-09-02", "turno_requerido": "COME"},
        {"fecha": "2025-09-03", "turno_requerido": "COME"},
        {"fecha": "2025-09-04", "turno_requerido": "COME"},
        {"fecha": "2025-09-05", "turno_requerido": "COME"},
    ],
    "YIS": [
        {"fecha": "2025-09-08", "turno_requerido": "COME"},
        {"fecha": "2025-09-09", "turno_requerido": "COME"},
        {"fecha": "2025-09-10", "turno_requerido": "COME"},
        {"fecha": "2025-09-11", "turno_requerido": "COME"},
        {"fecha": "2025-09-12", "turno_requerido": "COME"},
    ],
}

# ============================================================================
# TURNOS ESPECIALES
# ============================================================================

TURNOS_ESPECIALES = {
    "GMT": [
        {
            "tipo": "SIND",
            "frecuencia": "semanal_fijo",
            "dia_semana": "miércoles"
        }
    ],
    "GCE": [
        {
            "tipo": "SIND",
            "frecuencia": "semanal_fijo", 
            "dia_semana": "miércoles"
        }
    ]
}

# ============================================================================
# TRABAJADORES FUERA DE OPERACIÓN
# ============================================================================

TRABAJADORES_FUERA_OPERACION = ['PHD', 'WEH', 'VCM', 'MEI', 'ROP']

# ============================================================================
# DÍAS FESTIVOS
# ============================================================================

DIAS_FESTIVOS = [
    "2025-01-01", "2025-01-06", "2025-01-20", "2025-03-24", "2025-03-27", 
    "2025-03-28", "2025-03-30", "2025-05-01", "2025-05-12", "2025-06-02", 
    "2025-06-23", "2025-06-30", "2025-07-20", "2025-08-07", "2025-08-18", 
    "2025-10-13", "2025-11-03", "2025-11-17", "2025-12-08", "2025-12-25"
]

# ============================================================================
# CONFIGURACIÓN GENERAL
# ============================================================================

CONFIGURACION_GENERAL = {
    "archivo_historial_sabados": "historial_sabados.csv",
    "mapeo_dias": {
        "lunes": 0, "martes": 1, "miércoles": 2, "jueves": 3, 
        "viernes": 4, "sábado": 5, "domingo": 6
    },
    "turnos_validos": [
        # Turnos básicos
        "DESC", "TROP", 
        # Turnos completos
        "VACA", "COME", "COMT", "COMS",
        # Turnos adicionales originales
        "SIND", "CMED", "CERT",
        # Formación, instrucción y entrenamiento
        "CAPA", "MCAE", "TCAE", "MCHC", "TCHC", "NCHC", "ACHC",
        "MENT", "TENT", "NENT", "AENT",
        "MINS", "TINS", "NINS", "AINS",
        # Gestión, oficinas y grupos de trabajo
        "MCOR", "TCOR", "MSMS", "TSMS", "MDBM", "TDBM",
        "MDOC", "TDOC", "MPRO", "TPRO", "MATF", "TATF",
        "MGST", "TGST", "MOFI", "TOFI",
        # Operativos y asignaciones especiales
        "CET", "ATC", "KATC", "XATC", "YATC", "ZATC", "X"
    ],
    "turnos_completos": ["VACA", "COME", "COMT", "COMS"],
    "turnos_adicionales": [
        # Turnos adicionales originales
        "SIND", "CMED", "CERT",
        # Formación, instrucción y entrenamiento
        "CAPA", "MCAE", "TCAE", "MCHC", "TCHC", "NCHC", "ACHC",
        "MENT", "TENT", "NENT", "AENT",
        "MINS", "TINS", "NINS", "AINS",
        # Gestión, oficinas y grupos de trabajo
        "MCOR", "TCOR", "MSMS", "TSMS", "MDBM", "TDBM",
        "MDOC", "TDOC", "MPRO", "TPRO", "MATF", "TATF",
        "MGST", "TGST", "MOFI", "TOFI",
        # Operativos y asignaciones especiales
        "CET", "ATC", "KATC", "XATC", "YATC", "ZATC"
    ]
}

# ============================================================================
# LISTA DE EMPLEADOS
# ============================================================================

def obtener_empleados():
    """
    Genera la lista completa de empleados del sistema
    
    Returns:
        list: Lista de códigos de empleados (25 trabajadores)
    """
    return [
        'PHD', 'HLG', 'MEI', 'VCM', 'ROP', 'ECE', 'WEH', 'DFB', 'MLS', 'FCE',
        'JBV', 'GMT', 'BRS', 'HZG', 'JIS', 'CDT', 'WGG', 'GCE', 'YIS', 'MAQ',
        'DJO', 'AFG', 'JLF', 'JMV'
    ]