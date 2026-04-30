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

# ===== 1. CONFIGURACIÓN DE PÁGINA =====
st.set_page_config(page_title="CRM Capacidad Instalada", layout="wide")

# ===== 2. CONFIGURACIÓN DE NEGOCIO =====
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
HORAS_DIA_LABORAL = 6[cite: 1]

TAREAS_DISPONIBLE_TIPO = [
    "DISPONIBLE",
    "PLANIFICACIONES/ORGANIZACIÓN/PROCEDIMIENTO S/INFORMES",
    "INASISTENCIA POR EXAMEN O TRAMITE"
]

MESES_ES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
    7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
}

# ===== 3. CARGA DE FERIADOS =====
@st.cache_data
def cargar_feriados_desde_csv():
    try:
        df_f = pd.read_csv("feriados_2026.csv")
        df_f.columns = df_f.columns.str.strip().str.lower()
        if 'fecha' in df_f.columns:
            df_f['fecha'] = pd.to_datetime(df_f['fecha'], errors='coerce')[cite: 1]
            return df_f['fecha'].dt.date.dropna().tolist()
        return []
    except Exception:
        return []

FERIADOS = cargar_feriados_desde_csv()

# ===== 4. FUNCIONES DE PDF =====
def generar_pdf_reporte(nombre, mes, anio, total_hs, eficiencia, estado, df_tareas):
    from reportlab.lib.pagesizes import letter
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib import colors
    
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter)
    styles = getSampleStyleSheet()
    story = []
    
    story.append(Paragraph(f"Reporte de Capacidad - {nombre}", styles['Title']))
    story.append(Paragraph(f"Período: {MESES_ES[mes]} {anio}", styles['Normal']))
    story.append(Spacer(1, 12))
    
    data = [
        ["Métrica", "Valor"],
        ["Total Horas Cargadas", f"{total_hs} hs"],
        ["Eficiencia Operativa", f"{eficiencia:.1f}%"],
        ["Estado", estado]
    ]
    t = Table(data, colWidths=[200, 200])
    t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0), colors.grey), ('GRID',(0,0),(-1,-1),1,colors.black)]))
    story.append(t)
    
    story.append(Spacer(1, 20))
    story.append(Paragraph("Detalle por Tarea", styles['Heading2']))
    
    data_t = [["Tarea", "Horas"]]
    for _, row in df_tareas.iterrows():
        data_t.append([row['Tarea'], f"{row['Horas']} hs"])
    
    t_det = Table(data_t, colWidths=[300, 100])
    t_det.setStyle(TableStyle([('GRID',(0,0),(-1,-1),1,colors.black)]))
    story.append(t_det)
    
    doc.build(story)
    buffer.seek(0)
    return buffer

# ===== 5. CONEXIÓN GOOGLE SHEETS =====
@st.cache_resource
def conectar_sheets():
    scopes = ["https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scopes)
    return gspread.authorize(creds).open("CRM_Estudio_Datos")

@st.cache_data(ttl=60)
def cargar_hoja(nombre_hoja):
    try:
        ws = conectar_sheets().worksheet(nombre_hoja)
        df = pd.DataFrame(ws.get_all_records())
        for col in df.columns:
            if 'fecha' in col.lower():
                df[col] = pd.to_datetime(df[col], dayfirst=True, errors='coerce')[cite: 1]
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
                df_copy[col] = pd.to_datetime(df_copy[col], errors='coerce').dt.strftime('%d/%m/%Y')[cite: 1]
        ws.update([df_copy.columns.values.tolist()] + df_copy.fillna('').astype(str).values.tolist())
        st.cache_data.clear()
        return True
    except: return False

# ===== 6. FUNCIONES CAPACIDAD =====
def calcular_dias_habiles(fecha_inicio, fecha_fin):
    return len(pd.bdate_range(start=fecha_inicio, end=fecha_fin, freq='C', holidays=FERIADOS))[cite: 1]

def calcular_capacidad_mensual(anio, mes):
    inicio = datetime(anio, mes, 1)
    fin = (datetime(anio, mes + 1, 1) if mes < 12 else datetime(anio + 1, 1, 1)) - timedelta(days=1)
    dias = calcular_dias_habiles(inicio.date(), fin.date())
    return dias, dias * HORAS_DIA_LABORAL

# ===== 7. LOGIN & SESSION STATE =====
if 'cargas' not in st.session_state:
    df, _ = cargar_hoja("Cargas")
    st.session_state.cargas = df if not df.empty else pd.DataFrame(columns=['Fecha', 'Tarea'] + OPERARIOS_FIJOS + ['Nota'])

if 'usuario_actual' not in st.session_state:
    st.session_state.usuario_actual = None

if st.session_state.usuario_actual is None:
    st.title("CRM Estudio Contable")
    user = st.selectbox("Usuario:", ["Seleccionar..."] + OPERARIOS_FIJOS + ["Admin - Ver todo"])
    if st.button("Entrar") and user != "Seleccionar...":
        st.session_state.usuario_actual = user
        st.rerun()
    st.stop()

# ===== 8. MENU =====
menu = st.sidebar.radio("Menú", ["Panel de Control", "Cargar Horas", "Resetear Datos"] if st.session_state.usuario_actual == "Admin - Ver todo" else ["Panel de Control", "Cargar Mis Horas"])

# ===== 9. PANEL DE CONTROL =====
if menu == "Panel de Control":
    st.title("Panel de Control - Ocupación")
    c_a, c_m = st.columns(2)
    with c_a: anio = st.selectbox("Año", [2025, 2026, 2027], index=1)
    with c_m: mes = st.selectbox("Mes", list(range(1,13)), format_func=lambda x: MESES_ES[x], index=datetime.now().month-1)

    dias_h, cap_base = calcular_capacidad_mensual(anio, mes)
    st.info(f"**{MESES_ES[mes]} {anio}**: {dias_h} días hábiles. Capacidad teórica: {cap_base} hs.")

    df_p = st.session_state.cargas.copy()
    if not df_p.empty:
        df_p['Fecha'] = pd.to_datetime(df_p['Fecha'], errors='coerce')
        df_mes = df_p[(df_p['Fecha'].dt.month == mes) & (df_p['Fecha'].dt.year == anio)]
        
        if not df_mes.empty:
            df_melt = df_mes.melt(id_vars=['Fecha', 'Tarea'], value_vars=OPERARIOS_FIJOS, var_name='Operario', value_name='Horas')
            df_melt['Horas'] = pd.to_numeric(df_melt['Horas'], errors='coerce').fillna(0).round(2)
            
            p_sel = st.selectbox("Operario:", OPERARIOS_FIJOS) if st.session_state.usuario_actual == "Admin - Ver todo" else st.session_state.usuario_actual
            df_ind = df_melt[df_melt['Operario'] == p_sel]
            
            res_t = df_ind.groupby('Tarea')['Horas'].sum().round(1).reset_index()[cite: 1]
            total_c = res_t['Horas'].sum().round(1)
            
            hs_prod = df_ind[~df_ind['Tarea'].isin(TAREAS_DISPONIBLE_TIPO)]['Horas'].sum().round(1)
            eficiencia = (hs_prod / cap_base * 100) if cap_base > 0 else 0[cite: 1]
            
            hs_libres = df_ind[df_ind['Tarea'].isin(TAREAS_DISPONIBLE_TIPO)]['Horas'].sum().round(1)
            porc_libres = (hs_libres / total_c * 100) if total_c > 0 else 0
            
            if porc_libres > 20: est = "🟢 OK"
            elif porc_libres >= 10: est = "🟡 Atención"
            else: est = "🔴 Al límite"

            col_g, col_m = st.columns([2,1])
            with col_g:
                fig = px.pie(res_t, values='Horas', names='Tarea', color='Tarea', color_discrete_map=COLORES_TAREAS)
                fig.update_layout(title=f"Distribución {p_sel} ({total_c} hs)")
                st.plotly_chart(fig, use_container_width=True)
                
            with col_m:
                st.metric("Total Cargado", f"{total_c} hs")
                st.metric("Eficiencia Operativa", f"{eficiencia:.1f}%")[cite: 1]
                st.metric("Disponibilidad", f"{porc_libres:.1f}%")
                st.subheader(f"Estado: {est}")
                
                pdf_b = generar_pdf_reporte(p_sel, mes, anio, total_c, eficiencia, est, res_t)
                st.download_button("📄 Descargar PDF", data=pdf_b, file_name=f"Reporte_{p_sel}_{mes}.pdf")[cite: 1]

# ===== 10. CARGAR HORAS =====
elif "Cargar" in menu:
    st.title("Cargar Horas")
    u = st.selectbox("Persona:", OPERARIOS_FIJOS) if st.session_state.usuario_actual == "Admin - Ver todo" else st.session_state.usuario_actual
    with st.form("f"):
        f_f = st.date_input("Fecha", value=datetime.now())
        f_t = st.selectbox("Tarea", list(COLORES_TAREAS.keys()))
        f_h = st.number_input("Horas", step=0.5)
        if st.form_submit_button("Guardar"):
            nueva = {'Fecha': f_f, 'Tarea': f_t}
            for op in OPERARIOS_FIJOS: nueva[op] = round(f_h, 2) if op == u else 0[cite: 1]
            st.session_state.cargas = pd.concat([st.session_state.cargas, pd.DataFrame([nueva])], ignore_index=True)
            guardar_df("Cargas", st.session_state.cargas)
            st.success("Guardado")
            st.rerun()

elif menu == "Resetear Datos":
    if st.text_input("Escriba BORRAR") == "BORRAR":
        if st.button("Eliminar"):
            guardar_df("Cargas", pd.DataFrame(columns=['Fecha', 'Tarea'] + OPERARIOS_FIJOS))
            st.rerun()
