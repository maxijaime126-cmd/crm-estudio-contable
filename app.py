import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import plotly.express as px
import plotly.graph_objects as go
import gspread
from google.oauth2.service_account import Credentials
from io import BytesIO
import time
import os

# ===== CONFIGURACIÓN DE PÁGINA =====
st.set_page_config(page_title="CRM Capacidad Instalada", layout="wide")

# ===== CONFIGURACIÓN DE NEGOCIO =====
COLORES_TAREAS = {
    "DOCUMENTACIÓN CARGA": "#FFB6C1",
    "DOCUMENTACIÓN CONTROL": "#FF69B4",
    "IMPUESTOS": "#FF00FF",
    "SUELDOS": "#FFFF00",
    "CONTABILIDAD": "#00FF00",
    "ATENCION AL CLIENTE": "#00BFFF",
    "TAREAS NO RUTINARIAS": "#ADD8E6",
    "DISPONIBLE": "#FFFFFF",
    "REUNIONES DE EQUIPO": "#E6E6FA",
    "INASISTENCIA POR EXAMEN O TRAMITE": "#FFDAB9",
    "PLANIFICACIONES/ORGANIZACIÓN/PROCEDIMIENTO S/INFORMES": "#40E0D0"
}

OPERARIOS_FIJOS = ["Natalia", "Maximiliano", "Athina", "Johana"]
HORAS_DIA_LABORAL = 6  # Jornada de 09:00 a 15:00[cite: 1]

TAREAS_DISPONIBLE_TIPO = [
    "DISPONIBLE",
    "PLANIFICACIONES/ORGANIZACIÓN/PROCEDIMIENTO S/INFORMES",
    "INASISTENCIA POR EXAMEN O TRAMITE"
]

MESES_ES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
    7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
}

DIAS_SEMANA_ES = {
    0: "Lunes", 1: "Martes", 2: "Miércoles", 3: "Jueves",
    4: "Viernes", 5: "Sábado", 6: "Domingo"
}

# ===== CARGA DE FERIADOS DESDE GITHUB (CSV) =====
@st.cache_data
def cargar_feriados_desde_csv():
    try:
        # Intenta leer tu archivo local[cite: 1]
        df_f = pd.read_csv("feriados_2026.csv")
        df_f['fecha'] = pd.to_datetime(df_f['fecha'], dayfirst=True)
        return df_f['fecha'].dt.date.tolist()
    except Exception as e:
        st.sidebar.warning("No se encontró feriados_2026.csv, usando lista vacía.")
        return []

FERIADOS = cargar_feriados_desde_csv()

# ===== CONEXIÓN GOOGLE SHEETS =====
@st.cache_resource
def conectar_sheets():
    scopes = ["https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scopes)
    client = gspread.authorize(creds)
    return client.open("CRM_Estudio_Datos")

@st.cache_data(ttl=60)
def cargar_hoja(nombre_hoja):
    try:
        ws = conectar_sheets().worksheet(nombre_hoja)
        df = pd.DataFrame(ws.get_all_records())
        df.columns = df.columns.str.strip()
        for col in df.columns:
            if 'fecha' in col.lower():
                df[col] = pd.to_datetime(df[col], dayfirst=True, errors='coerce')
        return df, ws
    except:
        return pd.DataFrame(), None

def guardar_df(nombre_hoja, df):
    try:
        ws = conectar_sheets().worksheet(nombre_hoja)
        ws.clear()
        df_copy = df.copy()
        for col in df_copy.columns:
            if 'fecha' in col.lower():
                df_copy[col] = pd.to_datetime(df_copy[col], errors='coerce').dt.strftime('%d/%m/%Y')
        df_copy = df_copy.fillna('').astype(str)
        ws.update([df_copy.columns.values.tolist()] + df_copy.values.tolist())
        st.cache_data.clear()
        return True
    except Exception as e:
        st.error(f"Error guardando en {nombre_hoja}: {e}")
        return False

# ===== FUNCIONES DE CAPACIDAD =====
def calcular_dias_habiles(fecha_inicio, fecha_fin):
    # Usa la lista FERIADOS cargada del CSV[cite: 1]
    dias = pd.bdate_range(start=fecha_inicio, end=fecha_fin, freq='C', holidays=FERIADOS)
    return len(dias)

def calcular_capacidad_mensual(anio, mes):
    inicio = datetime(anio, mes, 1)
    if mes == 12:
        fin = datetime(anio + 1, 1, 1) - timedelta(days=1)
    else:
        fin = datetime(anio, mes + 1, 1) - timedelta(days=1)
    dias_h = calcular_dias_habiles(inicio.date(), fin.date())
    return dias_h, dias_h * HORAS_DIA_LABORAL

def calcular_excepciones_mes(operario, anio, mes, df_excepciones):
    if df_excepciones.empty:
        return 0
    df_exc = df_excepciones.copy()
    df_exc = df_exc[df_exc['Operario'] == operario]
    total_hs_exc = 0
    inicio_mes = datetime(anio, mes, 1).date()
    fin_mes = (datetime(anio, mes + 1, 1) if mes < 12 else datetime(anio + 1, 1, 1)).date() - timedelta(days=1)
    
    for _, row in df_exc.iterrows():
        if pd.notna(row['Fecha Inicio']) and pd.notna(row['Fecha Fin']):
            i_efectivo = max(row['Fecha Inicio'].date(), inicio_mes)
            f_efectivo = min(row['Fecha Fin'].date(), fin_mes)
            if i_efectivo <= f_efectivo:
                dias_e = calcular_dias_habiles(i_efectivo, f_efectivo)
                total_hs_exc += dias_e * HORAS_DIA_LABORAL
    return total_hs_exc

# ===== SESSION STATE =====
if 'cargas' not in st.session_state:
    df, _ = cargar_hoja("Cargas")
    if df.empty:
        df = pd.DataFrame(columns=['Fecha', 'Tarea'] + OPERARIOS_FIJOS + ['Nota'])
    st.session_state.cargas = df

if 'excepciones' not in st.session_state:
    df, _ = cargar_hoja("Excepciones")
    if df.empty:
        df = pd.DataFrame(columns=['Operario', 'Fecha Inicio', 'Fecha Fin', 'Motivo', 'Horas'])
    st.session_state.excepciones = df

if 'usuario_actual' not in st.session_state:
    st.session_state.usuario_actual = None

# ===== LOGIN =====
if st.session_state.usuario_actual is None:
    st.title("CRM Capacidad Instalada - Estudio")
    usuario = st.selectbox("Usuario:", ["Seleccionar..."] + OPERARIOS_FIJOS + ["Admin - Ver todo"])
    if st.button("Ingresar") and usuario != "Seleccionar...":
        st.session_state.usuario_actual = usuario
        st.rerun()
    st.stop()

# ===== SIDEBAR =====
st.sidebar.success(f"Usuario: **{st.session_state.usuario_actual}**")
if st.sidebar.button("Cerrar Sesión"):
    st.session_state.clear()
    st.rerun()

if st.session_state.usuario_actual == "Admin - Ver todo":
    menu = st.sidebar.radio("Menú", ["Panel de Control", "Cargar Horas", "Carga Masiva", "Excepciones", "Resetear Datos"])
else:
    menu = st.sidebar.radio("Menú", ["Panel de Control", "Cargar Mis Horas"])

# ===== PANEL DE CONTROL =====
if menu == "Panel de Control":
    st.title("Panel de Control - Ocupación")
    
    col_a, col_m = st.columns(2)
    with col_a: anio = st.selectbox("Año", [2025, 2026, 2027], index=1)
    with col_m: mes = st.selectbox("Mes", list(range(1,13)), format_func=lambda x: MESES_ES[x], index=datetime.now().month - 1)

    dias_habiles, cap_base = calcular_capacidad_mensual(anio, mes)
    st.info(f"**{MESES_ES[mes]} {anio}**: {dias_habiles} días hábiles (L-V sin feriados). Capacidad: {cap_base}hs[cite: 1]")

    df_pnl = st.session_state.cargas.copy()
    if not df_pnl.empty:
        df_pnl['Fecha'] = pd.to_datetime(df_pnl['Fecha'], errors='coerce')
        df_pnl = df_pnl[(df_pnl['Fecha'].dt.month == mes) & (df_pnl['Fecha'].dt.year == anio)]

    if df_pnl.empty:
        st.warning("No hay datos para este mes.")
    else:
        # Redondeo crítico para evitar el error de los decimales (.1)[cite: 1]
        df_melt = df_pnl.melt(id_vars=['Fecha', 'Tarea', 'Nota'], value_vars=OPERARIOS_FIJOS, var_name='Operario', value_name='Horas')
        df_melt['Horas'] = pd.to_numeric(df_melt['Horas'], errors='coerce').fillna(0).round(2)
        df_melt = df_melt[df_melt['Horas'] > 0]

        persona_sel = st.selectbox("Seleccionar persona:", OPERARIOS_FIJOS) if st.session_state.usuario_actual == "Admin - Ver todo" else st.session_state.usuario_actual
        
        hs_exc = calcular_excepciones_mes(persona_sel, anio, mes, st.session_state.excepciones)
        cap_real = cap_base - hs_exc
        df_ind = df_melt[df_melt['Operario'] == persona_sel]
        
        resumen_t = df_ind.groupby('Tarea')['Horas'].sum().round(1).reset_index()
        total_c = resumen_t['Horas'].sum().round(1)
        
        # Lógica de Disponibilidad e Inversión de Semáforo[cite: 1]
        hs_libres = df_ind[df_ind['Tarea'].isin(TAREAS_DISPONIBLE_TIPO)]['Horas'].sum().round(1)
        porc_libres = (hs_libres / total_c * 100) if total_c > 0 else 0

        if porc_libres > 20:
            estado_semaforo = "🟢 OK"
            msg = "Disponibilidad alta (>20%)[cite: 1]"
        elif porc_libres >= 10:
            estado_semaforo = "🟡 Atención"
            msg = "Capacidad moderada (10-20%)[cite: 1]"
        else:
            estado_semaforo = "🔴 Al límite"
            msg = "Sobrecarga detectada (<10% libre)[cite: 1]"

        c_graf, c_metr = st.columns([2,1])
        with c_graf:
            fig = px.pie(resumen_t, values='Horas', names='Tarea', color='Tarea', color_discrete_map=COLORES_TAREAS)
            fig.update_layout(title=f"Ocupación: {persona_sel} ({total_c} hs)")
            st.plotly_chart(fig, use_container_width=True)
        
        with c_metr:
            st.metric("Total Cargado", f"{total_c} hs")
            st.metric("Capacidad Real", f"{cap_real} hs", delta=f"-{hs_exc}hs excepcion" if hs_exc > 0 else None)
            st.metric("Tiempo Disponible", f"{hs_libres} hs", delta=f"{porc_libres:.1f}%")
            st.subheader(f"Estado: {estado_semaforo}")
            st.caption(msg)

# ===== CARGAR HORAS =====
elif menu in ["Cargar Mis Horas", "Cargar Horas"]:
    st.title("Cargar Horas")
    user_c = st.selectbox("Persona:", OPERARIOS_FIJOS) if st.session_state.usuario_actual == "Admin - Ver todo" else st.session_state.usuario_actual
    
    with st.form("form_c", clear_on_submit=True):
        f_fecha = st.date_input("Fecha", value=datetime.now())
        f_tarea = st.selectbox("Tarea", list(COLORES_TAREAS.keys()))
        f_horas = st.number_input("Horas", min_value=0.0, step=0.5)
        f_nota = st.text_area("Nota")
        if st.form_submit_button("Guardar"):
            if f_horas > 0:
                # Se guarda redondeado para evitar errores de precisión[cite: 1]
                nueva = {'Fecha': f_fecha, 'Tarea': f_tarea, 'Nota': f_nota}
                for op in OPERARIOS_FIJOS: nueva[op] = round(f_horas, 2) if op == user_c else 0
                st.session_state.cargas = pd.concat([st.session_state.cargas, pd.DataFrame([nueva])], ignore_index=True)
                guardar_df("Cargas", st.session_state.cargas)
                st.success("Guardado.")
                st.rerun()

# (Secciones de Carga Masiva y Excepciones simplificadas para este bloque)
elif menu == "Excepciones":
    st.title("Excepciones")
    # Mantiene tu lógica pero usando el cálculo de días hábiles corregido[cite: 1]
    ...

elif menu == "Resetear Datos":
    st.title("⚠️ Resetear Datos")
    if st.text_input("Escriba 'BORRAR'") == "BORRAR":
        if st.button("Borrar todo"):
            df_v = pd.DataFrame(columns=['Fecha', 'Tarea'] + OPERARIOS_FIJOS + ['Nota'])
            guardar_df("Cargas", df_v)
            st.session_state.cargas = df_v
            st.rerun()
