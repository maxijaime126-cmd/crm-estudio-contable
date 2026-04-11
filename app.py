import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

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
def cargar_cargas():
    sheet = conectar_sheets()
    worksheet = sheet.worksheet("Cargas")
    data = worksheet.get_all_records()
    df = pd.DataFrame(data)
    
    if df.empty:
        return df
    
    df.columns = df.columns.str.strip()
    
    if 'Fecha' in df.columns:
        df['Fecha'] = pd.to_datetime(df['Fecha'], dayfirst=True, errors='coerce').dt.date
    
    return df

# ===== INTERFAZ =====
st.title("📊 CRM Estudio Contable")

# Selector de usuario
usuarios = ["Maximiliano", "Natalia", "Athina", "Johana"]
usuario_carga = st.sidebar.selectbox("Seleccionar usuario:", usuarios)

st.header(f"Tareas de {usuario_carga}")

# Cargamos datos
try:
    cargas_df = cargar_cargas()
except Exception as e:
    st.error(f"Error al conectar con Google Sheets: {e}")
    st.info("Verificá que: 1) La pestaña se llame 'Cargas' 2) Esté compartido con bot-564@watchful-gear-493001-m5.iam.gserviceaccount.com")
    st.stop()

if cargas_df.empty:
    st.warning("El Google Sheet está vacío. Cargá la primera fila con columnas: Fecha, Tarea, Usuario, Nota")
    st.stop()

if 'Usuario' not in cargas_df.columns:
    st.error("Tu Sheet no tiene la columna 'Usuario'. Las columnas deben ser: Fecha | Tarea | Usuario | Nota")
    st.dataframe(cargas_df)
    st.stop()

# Filtramos por usuario
df_mis_cargas = cargas_df[cargas_df['Usuario'] == usuario_carga].copy()

if df_mis_cargas.empty:
    st.info(f"No tenés tareas asignadas, {usuario_carga}")
else:
    st.dataframe(df_mis_cargas, use_container_width=True, hide_index=True)
    
    # Descargar CSV
    csv = df_mis_cargas.to_csv(index=False).encode('utf-8')
    st.download_button(
        "⬇️ Descargar mis tareas", 
        csv, 
        f"tareas_{usuario_carga}.csv", 
        "text/csv"
    )

# ===== MOSTRAR TODO EL SHEET =====
with st.expander("Ver todas las cargas"):
    st.dataframe(cargas_df, use_container_width=True, hide_index=True)
