import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, date
import numpy as np

# ===== CONFIGURACIÓN =====
SCOPE = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

# Leemos las credenciales desde st.secrets
creds = Credentials.from_service_account_info(
    st.secrets["gcp_service_account"],
    scopes=SCOPE
)
client = gspread.authorize(creds)

# Abrimos el Sheet por nombre
SHEET_NAME = "CRM_Estudio_Datos"
sh = client.open(SHEET_NAME)

USUARIOS = ["Admin", "Natalia", "Lautaro", "Maxi"]
CLIENTES = ["Estudio Contable A", "Cliente B", "Cliente C"]
TAREAS = ["Liquidación", "IVA", "Ganancias", "Sueldos", "Monotributo", "Consulta", "Otro"]

# ===== FUNCIONES =====
def cargar_df(nombre_hoja):
    """Lee una hoja y devuelve DataFrame. Si está vacía, devuelve df con columnas correctas"""
    try:
        ws = sh.worksheet(nombre_hoja)
        data = ws.get_all_records()
        if not data:
            if nombre_hoja == "Cargas":
                return pd.DataFrame(columns=["Fecha", "Usuario", "Cliente", "Tarea", "Horas", "Comentario", "Timestamp_Carga"])
            else:
                return pd.DataFrame()
        df = pd.DataFrame(data)
        return df
    except gspread.exceptions.WorksheetNotFound:
        if nombre_hoja == "Cargas":
            return pd.DataFrame(columns=["Fecha", "Usuario", "Cliente", "Tarea", "Horas", "Comentario", "Timestamp_Carga"])
        return pd.DataFrame()

def guardar_df(nombre_hoja, df):
    """Guarda el DataFrame en Google Sheets"""
    ws = sh.worksheet(nombre_hoja)
    ws.clear()

    df_copy = df.copy()

    # FIX: Convertir fechas y NaN para que gspread no rompa
    for col in df_copy.columns:
        if pd.api.types.is_datetime64_any_dtype(df_copy[col]):
            df_copy[col] = df_copy[col].dt.strftime('%Y-%m-%d')

    df_copy = df_copy.fillna('').astype(str)

    ws.update([df_copy.columns.values.tolist()] + df_copy.values.tolist())

# ===== LOGIN SIMPLE =====
st.set_page_config(page_title="CRM Estudio", page_icon="⏱️", layout="wide")

if "usuario" not in st.session_state:
    st.session_state.usuario = None

if st.session_state.usuario is None:
    st.title("⏱️ CRM Estudio Contable")
    st.subheader("Seleccioná tu usuario para ingresar")
    usuario_seleccionado = st.selectbox("Usuario", USUARIOS)
    if st.button("Ingresar", type="primary"):
        st.session_state.usuario = usuario_seleccionado
        st.rerun()
    st.stop()

# ===== APP PRINCIPAL =====
usuario_actual = st.session_state.usuario
es_admin = usuario_actual == "Admin"

st.sidebar.success(f"Usuario: {usuario_actual}")
if st.sidebar.button("Cerrar sesión"):
    st.session_state.usuario = None
    st.rerun()

st.title("⏱️ Carga de Horas")

# Cargamos datos
if "cargas" not in st.session_state:
    st.session_state.cargas = cargar_df("Cargas")

# ===== FORMULARIO DE CARGA =====
with st.form("form_carga", clear_on_submit=True):
    col1, col2, col3 = st.columns(3)
    with col1:
        fecha = st.date_input("Fecha", value=date.today())
    with col2:
        cliente = st.selectbox("Cliente", CLIENTES)
    with col3:
        tarea = st.selectbox("Tarea", TAREAS)

    col4, col5 = st.columns([1, 3])
    with col4:
        horas = st.number_input("Horas", min_value=0.0, max_value=24.0, step=0.25)
    with col5:
        comentario = st.text_input("Comentario", placeholder="Opcional")

    submitted = st.form_submit_button("💾 Guardar horas", type="primary")

    if submitted:
        if horas <= 0:
            st.error("Las horas deben ser mayor a 0")
        else:
            nueva_fila = {
                "Fecha": fecha,
                "Usuario": usuario_actual,
                "Cliente": cliente,
                "Tarea": tarea,
                "Horas": horas,
                "Comentario": comentario,
                "Timestamp_Carga": datetime.now()
            }
            st.session_state.cargas = pd.concat(
                [st.session_state.cargas, pd.DataFrame([nueva_fila])],
                ignore_index=True
            )
            guardar_df("Cargas", st.session_state.cargas)
            st.success(f"✅ Guardado: {horas} hs en {cliente}")
            st.rerun()

# ===== TABLA DE CARGAS =====
st.divider()
st.subheader("Mis horas cargadas")

df_mostrar = st.session_state.cargas.copy()

# Si no es admin, filtrar solo sus horas
if not es_admin:
    df_mostrar = df_mostrar[df_mostrar["Usuario"] == usuario_actual]

# Ordenar por fecha descendente
if not df_mostrar.empty:
    df_mostrar["Fecha"] = pd.to_datetime(df_mostrar["Fecha"])
    df_mostrar = df_mostrar.sort_values("Fecha", ascending=False)
    df_mostrar["Fecha"] = df_mostrar["Fecha"].dt.strftime('%Y-%m-%d')

st.dataframe(df_mostrar, use_container_width=True, hide_index=True)

# ===== DASHBOARD ADMIN =====
if es_admin:
    st.divider()
    st.subheader("📊 Dashboard Admin")

    if not st.session_state.cargas.empty:
        df_dash = st.session_state.cargas.copy()
        df_dash["Fecha"] = pd.to_datetime(df_dash["Fecha"])
        df_dash["Horas"] = pd.to_numeric(df_dash["Horas"])

        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total Horas Cargadas", f"{df_dash['Horas'].sum():.2f} hs")
        with col2:
            st.metric("Total Registros", len(df_dash))
        with col3:
            st.metric("Clientes Activos", df_dash['Cliente'].nunique())

        st.write("**Horas por Usuario**")
        st.bar_chart(df_dash.groupby("Usuario")["Horas"].sum())

        st.write("**Horas por Cliente**")
        st.bar_chart(df_dash.groupby("Cliente")["Horas"].sum())
    else:
        st.info("Todavía no hay horas cargadas.")
