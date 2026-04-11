import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import plotly.express as px
from datetime import date

st.set_page_config(page_title="CRM Estudio Contable", layout="wide")

# ===== CONEXIÓN A GOOGLE SHEETS =====
@st.cache_resource
def conectar_sheets():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=scopes
    )
    client = gspread.authorize(creds)
    sheet = client.open("CRM_Estudio_Datos")
    return sheet

@st.cache_data(ttl=60)
def cargar_datos(nombre_hoja):
    sheet = conectar_sheets()
    worksheet = sheet.worksheet(nombre_hoja)
    data = worksheet.get_all_records()
    df = pd.DataFrame(data)
    df.columns = df.columns.str.strip()
    return df

def guardar_horas(tarea, usuario, horas, nota):
    sheet = conectar_sheets()
    try:
        ws = sheet.worksheet("Horas_Cargadas")
    except:
        ws = sheet.add_worksheet(title="Horas_Cargadas", rows=1000, cols=10)
        ws.append_row(["Fecha_Carga", "Fecha_Tarea", "Tarea", "Usuario", "Horas", "Nota"])

    ws.append_row([
        str(date.today()),
        tarea.get('Fecha', ''),
        tarea.get('Tarea', ''),
        usuario,
        horas,
        nota
    ])
    st.cache_data.clear()

# ===== PANTALLA INICIAL =====
if 'usuario_actual' not in st.session_state:
    st.session_state.usuario_actual = None

if st.session_state.usuario_actual is None:
    st.title("📊 CRM Estudio Contable")
    st.subheader("¿Quién va a cargar?")

    col1, col2, col3, col4, col5 = st.columns(5)
    usuarios = ["Maximiliano", "Natalia", "Athina", "Johana", "Administrador"]

    for i, user in enumerate(usuarios):
        with [col1, col2, col3, col4, col5][i]:
            if st.button(user, use_container_width=True, type="primary"):
                st.session_state.usuario_actual = user
                st.rerun()
    st.stop()

# ===== HEADER CON LOGOUT =====
usuario = st.session_state.usuario_actual
col1, col2 = st.columns([4,1])
with col1:
    st.title(f"CRM - {usuario}")
with col2:
    if st.button("Cambiar usuario"):
        st.session_state.usuario_actual = None
        st.rerun()

# ===== CARGAMOS DATOS =====
try:
    tareas_df = cargar_datos("Cargas")
    horas_df = cargar_datos("Horas_Cargadas")
except Exception as e:
    st.error(f"Error: {e}")
    st.info("Verificá que exista la pestaña 'Cargas' y esté compartido con el bot")
    st.stop()

# ===== VISTA USUARIO NORMAL =====
if usuario!= "Administrador":
    st.header("Mis tareas asignadas")
    mis_tareas = tareas_df[tareas_df['Usuario'] == usuario].copy()

    if mis_tareas.empty:
        st.info("No tenés tareas asignadas")
        st.stop()

    # Mostramos cada tarea para cargar horas
    for idx, tarea in mis_tareas.iterrows():
        with st.container(border=True):
            col1, col2 = st.columns([3,1])
            with col1:
                st.subheader(tarea['Tarea'])
                st.caption(f"Vence: {tarea['Fecha']} | Nota: {tarea.get('Nota', 'Sin nota')}")
            with col2:
                # Vemos si ya cargó horas
                ya_cargada = horas_df[
                    (horas_df['Usuario'] == usuario) &
                    (horas_df['Tarea'] == tarea['Tarea']) &
                    (horas_df['Fecha_Tarea'] == str(tarea['Fecha']))
                ]
                if not ya_cargada.empty:
                    st.success(f"Ya cargaste: {ya_cargada['Horas'].iloc[0]} hs")
                else:
                    with st.form(key=f"form_{idx}"):
                        horas = st.number_input("Horas que te llevó", min_value=0.0, step=0.5, key=f"h_{idx}")
                        nota = st.text_input("Nota", key=f"n_{idx}")
                        if st.form_submit_button("Guardar"):
                            guardar_horas(tarea, usuario, horas, nota)
                            st.success("Guardado!")
                            st.rerun()

# ===== VISTA ADMINISTRADOR =====
else:
    st.header("Panel Administrador")

    if horas_df.empty:
        st.warning("Todavía nadie cargó horas")
        st.stop()

    # Filtros
    usuarios_filtro = st.multiselect(
        "Filtrar por usuario",
        options=horas_df['Usuario'].unique(),
        default=horas_df['Usuario'].unique()
    )
    df_filtrado = horas_df[horas_df['Usuario'].isin(usuarios_filtro)]

    # Métricas
    col1, col2, col3 = st.columns(3)
    col1.metric("Total Horas Cargadas", f"{df_filtrado['Horas'].sum():.1f} hs")
    col2.metric("Tareas Completadas", len(df_filtrado))
    col3.metric("Usuarios Activos", df_filtrado['Usuario'].nunique())

    # Gráfico de torta
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Horas por Usuario")
        fig1 = px.pie(df_filtrado, names='Usuario', values='Horas', hole=0.3)
        st.plotly_chart(fig1, use_container_width=True)

    with col2:
        st.subheader("Total General vs Por Usuario")
        total_general = df_filtrado['Horas'].sum()
        data_torta = pd.DataFrame({
            'Tipo': ['Total General'] + list(df_filtrado['Usuario'].unique()),
            'Horas': [total_general] + [df_filtrado[df_filtrado['Usuario']==u]['Horas'].sum() for u in df_filtrado['Usuario'].unique()]
        })
        fig2 = px.pie(data_torta, names='Tipo', values='Horas', hole=0.3)
        st.plotly_chart(fig2, use_container_width=True)

    # Tabla detalle
    with st.expander("Ver detalle de cargas"):
        st.dataframe(df_filtrado, use_container_width=True, hide_index=True)
