"""Configuración central del CRM v2. Lo que cambia seguido (catálogo, personas,
feriados) vive en el Google Sheet, no acá."""

from datetime import date

SPREADSHEET = "CRM_Estudio_Datos"
TZ = "America/Argentina/Buenos_Aires"

# Desde esta fecha se reclaman los días sin cargar (el aviso de días pendientes no mira antes)
FECHA_INICIO = date(2026, 10, 1)

# --- Tiempo ---
HORAS_DIA_DEFAULT = 6.0
MIN_MINUTOS = 10        # mínimo por carga
PASO_MINUTOS = 5        # los minutos se cargan de a 5
MAX_MINUTOS_CARGA = 720  # tope por carga (12 hs), evita errores de tipeo

# --- Hojas del Sheet (las viejas "Cargas" y "Competencias" no se tocan) ---
HOJA_CATALOGO = "Catalogo"
HOJA_PERSONAS = "Personas"
HOJA_REGISTROS = "Registros"
HOJA_FERIADOS = "Feriados"

COLS_CATALOGO = ["Departamento", "Tarea", "Subtarea", "Tipo", "Orden", "Activo"]
COLS_PERSONAS = ["Nombre", "Rol", "Activo", "HorasDia"]
COLS_REGISTROS = ["ID", "Fecha", "Persona", "Departamento", "Tarea", "Subtarea",
                  "Minutos", "Nota", "Registrado"]
COLS_FERIADOS = ["Fecha", "Motivo"]

ESQUEMA = {
    HOJA_CATALOGO: COLS_CATALOGO,
    HOJA_PERSONAS: COLS_PERSONAS,
    HOJA_REGISTROS: COLS_REGISTROS,
    HOJA_FERIADOS: COLS_FERIADOS,
}
# Hojas que se completan con datos iniciales si están vacías
SEMBRAR_SI_VACIA = (HOJA_CATALOGO, HOJA_PERSONAS, HOJA_FERIADOS)

# --- Cuándo un día del equipo se considera cargado (trabajo / capacidad del equipo) ---
UMBRAL_ALTO = 60        # desde este % la carga se considera alta
UMBRAL_SATURADO = 85    # desde este % (o con 1 hora extra o más del equipo) el día se considera saturado

# --- Tipos de tarea (columna Tipo del catálogo) ---
TIPO_TRABAJO = "Trabajo"        # suma a la demanda
TIPO_DISPONIBLE = "Disponible"  # tiempo libre ("Disponible" y "Gestión y mejoras")
TIPO_AUSENCIA = "Ausencia"      # baja la capacidad (inasistencia, vacaciones)

TAREA_AUSENCIAS = "Ausencias"
TAREA_DISPONIBLE = "Disponible"
TAREA_GESTION = "Gestión y mejoras del departamento"
# Disponible y Ausencias no pertenecen a un departamento: se guardan con este nombre
DEPTO_GENERAL = "GENERAL"
SUBTAREA_NOTA_OBLIGATORIA = "Otros imprevistos"

# --- Personas iniciales (después se editan en la hoja Personas) ---
NOMBRE_ADMIN = "Admin - Ver todo"
ROL_ADMIN = "Admin"
ROL_OPERARIO = "Operario"
PERSONAS_SEED = [
    ("Natalia", ROL_OPERARIO, "SI", 6),
    ("Maximiliano", ROL_OPERARIO, "SI", 6),
    ("Athina", ROL_OPERARIO, "SI", 6),
    ("Johana", ROL_OPERARIO, "SI", 6),
    (NOMBRE_ADMIN, ROL_ADMIN, "SI", 0),
]

# --- Colores por departamento (los mismos del organigrama) ---
COLORES_DEPTO = {
    "SUELDOS": "#B8A200", "GERENCIAL": "#2A9D93", "TRÁMITES": "#C0504D",
    "IMPUESTOS": "#D6759F", "CONTABILIDAD": "#7F7F7F", "ATENCIÓN AL CLIENTE": "#3F7FBF",
    "RECURSOS HUMANOS": "#A8902B", "DOCUMENTACIÓN": "#3E9B3E", "COOPERATIVA": "#595959",
}

ICONOS_DEPTO = {
    "GENERAL": "🟦",
    "SUELDOS": "💰", "GERENCIAL": "🧭", "TRÁMITES": "⚖️", "IMPUESTOS": "🧾", "CONTABILIDAD": "📒",
    "ATENCIÓN AL CLIENTE": "📞", "RECURSOS HUMANOS": "🧑‍🤝‍🧑", "DOCUMENTACIÓN": "📄", "COOPERATIVA": "🤝",
}

MESES_ES = {1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
            7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre",
            12: "Diciembre"}
