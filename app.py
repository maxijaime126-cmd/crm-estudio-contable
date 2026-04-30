import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import plotly.express as px
import plotly.graph_objects as go
import gspread
from google.oauth2.service_account import Credentials
from io import BytesIO
import time

# ===== 1. CONFIGURACIÓN =====
st.set_page_config(page_title="CRM Capacidad Instalada", layout="wide")

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
HORAS_DIA_LABORAL = 6

TAREAS_DISPONIBLE_TIPO = ["DISPONIBLE", "PLANIFICACIONES/ORGANIZACIÓN/PROCEDIMIENTO S/INFORMES", "INASISTENCIA POR EXAMEN O TRAMITE"]

MESES_ES = {1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
            7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"}

DIAS_SEMANA_ES = {0: "Lunes", 1: "Martes", 2: "Miércoles", 3: "Jueves", 4: "Viernes", 5: "Sábado", 6: "Domingo"}

# ===== 2. FUNCIONES PDF (REPORTE Y PROTOCOLO ACTUALIZADO) =====
def generar_pdf_protocolo():
    from reportlab.lib.pagesizes import letter
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, ListFlowable, ListItem
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib.enums import TA_CENTER
    buf = BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=letter)
    s = getSampleStyleSheet()
    
    # Estilo personalizado para el Lema
    style_lema = getSampleStyleSheet()['Normal']
    style_lema.alignment = TA_CENTER
    style_lema.fontName = 'Helvetica-Oblique'

    story = [
        Paragraph("GRUPO PRESSACCO", s['Title']),
        Paragraph("<i>Elegimos sumar</i>", style_lema),
        Spacer(1, 20),
        Paragraph("Protocolo de Trabajo del Sistema", s['Heading1']),
        Spacer(1, 12),
        Paragraph("1. Responsabilidad Individual y Privacidad", s['Heading2']),
        Paragraph("Cada integrante debe ingresar con su usuario correspondiente. Es obligatorio cargar las horas únicamente en el panel personal para no distorsionar los datos de los demás compañeros y mantener la integridad de los reportes por sector.", s['Normal']),
        Spacer(1, 10),
        Paragraph("2. Reglas de Carga Diaria", s['Heading2']),
        ListFlowable([
            ListItem(Paragraph("Cargar horas todos los días antes de las 15:00 hs.", s['Normal'])),
            ListItem(Paragraph("La jornada debe completar siempre un total de 6 horas.", s['Normal'])),
            ListItem(Paragraph("Cargar bloques mínimos de 30 minutos (0.5 hs) para mayor precisión.", s['Normal'])),
            ListItem(Paragraph("Si no hay tareas operativas, registrar el tiempo restante como 'DISPONIBLE'.", s['Normal'])),
        ], bulletType='bullet'),
        Spacer(1, 10),
        Paragraph("3. Significado del Semáforo de Bienestar", s['Heading2']),
        Paragraph("🟢 VERDE: Más del 20% libre (Capacidad de sumar tareas).", s['Normal']),
        Paragraph("🟡 AMARILLO: 10% a 20% libre (Carga óptima).", s['Normal']),
        Paragraph("🔴 ROJO: Menos del 10% libre (Saturación / Necesidad de apoyo).", s['Normal']),
    ]
    doc.build(story)
    buf.seek(0)
    return buf

def generar_pdf_reporte(nombre, mes, anio, total_hs, eficiencia, estado, df_tareas):
    from reportlab.lib.pagesizes import letter
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib import colors
    buf = BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=letter)
    s = getSampleStyleSheet()
    story = [Paragraph(f"Reporte {nombre} - {MESES_ES[mes]} {anio}", s['Title']), Spacer(1,12)]
    data = [["Métrica", "Valor"], ["Total", f"{total_hs}hs"], ["Eficiencia", f"{eficiencia:.1f}%"], ["Estado", estado]]
    t = Table(data, colWidths=[200, 200])
    t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.grey),('GRID',(0,0),(-1,-1),1,colors.black)]))
    story.append(t)
    doc.build(story)
    buf.seek(0)
    return buf

# ===== 3. CONEXIÓN Y DATOS =====
@st.cache_resource
def conectar():
    creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=["https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/drive"])
    return gspread.authorize(creds).open("CRM_Estudio_Datos")

def cargar_hoja(nombre):
    try:
        ws = conectar().worksheet(nombre)
        df = pd.DataFrame(ws.get_all_records())
        for c in df.columns:
            if 'fecha' in c.lower(): df[c] = pd.to_datetime(df[c], dayfirst=True, errors='coerce')
        return df, ws
    except: return pd.DataFrame(), None

def guardar_df(nombre, df):
    try:
        ws = conectar().worksheet(nombre)
        ws.clear()
        df_c = df.copy()
        for c in df_c.columns:
            if 'fecha' in c.lower(): df_c[c] = pd.to_datetime(df_c[c], errors='coerce').dt.strftime('%d/%m/%Y')
        ws.update([df_c.columns.values.tolist()] + df_c.fillna('').astype(str).values.tolist())
        st.cache_data.clear()
        return True
    except: return False

@st.cache_data
def cargar_feriados():
    try:
        df_f = pd.read_csv("feriados_2026.csv")
        df_f.columns = df_f.columns.str.strip().str.lower()
        df_f['fecha'] = pd.to_datetime(df_f['fecha'], errors='coerce')
        return df_f['fecha'].dt.date.dropna().tolist()
    except: return []

FERIADOS = cargar_feriados()

# ===== 4. INICIO APP =====
if 'cargas' not in st.session_state:
    df, _ = cargar_hoja("Cargas")
    st.session_state.cargas = df if not df.empty else pd.DataFrame(columns=['Fecha', 'Tarea'] + OPERARIOS_FIJOS + ['Nota'])

if 'usuario_actual' not in st.session_state:
    st.session_state.usuario_actual = None

if st.session_state.usuario_actual is None:
    st.title("CRM Estudio Contable")
    u = st.selectbox("Usuario:", ["Seleccionar..."] + OPERARIOS_FIJOS + ["Admin - Ver todo"])
    if st.button("Ingresar") and u != "Seleccionar...":
        st.session_state.usuario_actual = u
        st.rerun()
    st.stop()

# ===== 5. MENÚ SIDEBAR (SOPORTE PARA ROLES) =====
opciones = ["Panel de Control", "Cargar Horas", "Carga Masiva", "Protocolo de Trabajo", "Resetear Datos"] if st.session_state.usuario_actual == "Admin - Ver todo" else ["Panel de Control", "Cargar Mis Horas", "Protocolo de Trabajo"]
menu = st.sidebar.radio("Menú", opciones)

if st.sidebar.button("Salir"):
    st.session_state.clear()
    st.rerun()

# ===== 6. SECCIONES =====

if menu == "Protocolo de Trabajo":
    st.title("📖 Protocolo: Grupo Pressacco")
    st.subheader("Elegimos sumar")
    st.info("Este protocolo garantiza que la información del estudio sea privada, organizada y útil para todos.")
    
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**1. Privacidad de los Datos**")
        st.write("Cada integrante debe seleccionar **únicamente su nombre** al cargar. Esto evita errores en los paneles de los demás compañeros.")
        st.markdown("**2. Horario de Carga**")
        st.write("Carga tus horas diariamente antes de las 15:00 hs. No dejes días sin completar.")
    with c2:
        st.markdown("**3. ¿Qué es el Semáforo?**")
        st.write("🟢 **OK:** Más del 20% libre.")
        st.write("🟡 **Atención:** 10% a 20% libre.")
        st.write("🔴 **Al límite:** Menos del 10% libre.")
    
    st.divider()
    st.download_button("📥 Descargar Protocolo Oficial (PDF)", data=generar_pdf_protocolo(), file_name="Protocolo_Grupo_Pressacco.pdf")

elif menu == "Panel de Control":
    st.title("Panel de Control - Ocupación")
    c1, c2 = st.columns(2)
    with c1: anio = st.selectbox("Año", [2025, 2026, 2027], index=1)
    with c2: mes = st.selectbox("Mes", list(range(1,13)), format_func=lambda x: MESES_ES[x], index=datetime.now().month-1)

    inicio_m = datetime(anio, mes, 1).date()
    fin_m = (datetime(anio, mes+1, 1) if mes < 12 else datetime(anio+1, 1, 1)).date() - timedelta(days=1)
    dias_h = len(pd.bdate_range(start=inicio_m, end=fin_m, freq='C', holidays=FERIADOS))
    cap_base = dias_h * HORAS_DIA_LABORAL
    st.info(f"**{MESES_ES[mes]} {anio}**: {dias_h} días hábiles. Capacidad: {cap_base} hs.")

    df_p = st.session_state.cargas.copy()
    if not df_p.empty:
        df_p['Fecha'] = pd.to_datetime(df_p['Fecha'], errors='coerce')
        df_m = df_p[(df_p['Fecha'].dt.month == mes) & (df_p['Fecha'].dt.year == anio)]
        
        p_sel = st.selectbox("Operario:", OPERARIOS_FIJOS) if st.session_state.usuario_actual == "Admin - Ver todo" else st.session_state.usuario_actual
        df_ind_melt = df_m.melt(id_vars=['Fecha', 'Tarea'], value_vars=[p_sel], var_name='Op', value_name='Hs')
        df_ind_melt = df_ind_melt[df_ind_melt['Hs'] > 0]
        
        res_t = df_ind_melt.groupby('Tarea')['Hs'].sum().round(1).reset_index()
        total_c = res_t['Hs'].sum().round(1)
        hs_p = df_ind_melt[~df_ind_melt['Tarea'].isin(TAREAS_DISPONIBLE_TIPO)]['Hs'].sum()
        eficiencia = (hs_p / cap_base * 100) if cap_base > 0 else 0
        porc_l = (df_ind_melt[df_ind_melt['Tarea'].isin(TAREAS_DISPONIBLE_TIPO)]['Hs'].sum() / total_c * 100) if total_c > 0 else 0
        
        if porc_l >= 20.0: est = "🟢 OK"
        elif porc_l >= 10.0: est = "🟡 Atención"
        else: est = "🔴 Al límite"

        col_g, col_m = st.columns([2,1])
        with col_g:
            st.plotly_chart(px.pie(res_t, values='Hs', names='Tarea', color='Tarea', color_discrete_map=COLORES_TAREAS, title=f"Ocupación {p_sel}"), use_container_width=True)
        with col_m:
            st.metric("Total Cargado", f"{total_c} hs")
            st.metric("Eficiencia", f"{eficiencia:.1f}%")
            st.metric("Disponibilidad", f"{porc_l:.1f}%")
            st.subheader(f"Estado: {est}")
            if st.button("Generar Reporte PDF"):
                pdf_r = generar_pdf_reporte(p_sel, mes, anio, total_c, eficiencia, est, res_t)
                st.download_button("Descargar Reporte", pdf_r, f"Reporte_{p_sel}.pdf")

elif menu == "Carga Masiva":
    st.title("Carga Masiva de Horas")
    with st.form("f_masiva"):
        u_m = st.selectbox("Operario", OPERARIOS_FIJOS)
        t_m = st.selectbox("Tarea", list(COLORES_TAREAS.keys()))
        f_i = st.date_input("Inicio"); f_f = st.date_input("Fin")
        h_t = st.number_input("Horas Totales", min_value=0.0)
        if st.form_submit_button("Guardar"):
            dias = pd.bdate_range(start=f_i, end=f_f, freq='C', holidays=FERIADOS)
            if len(dias) > 0:
                h_d = round(h_t / len(dias), 2)
                filas = []
                for d in dias:
                    f = {'Fecha': d, 'Tarea': t_m, 'Nota': 'Carga Masiva'}
                    for o in OPERARIOS_FIJOS: f[o] = h_d if o == u_m else 0
                    filas.append(f)
                st.session_state.cargas = pd.concat([st.session_state.cargas, pd.DataFrame(filas)], ignore_index=True)
                guardar_df("Cargas", st.session_state.cargas)
                st.success("Carga masiva completada"); time.sleep(1); st.rerun()

elif "Cargar" in menu:
    st.title("Cargar Horas")
    u_c = st.selectbox("Persona:", OPERARIOS_FIJOS) if st.session_state.usuario_actual == "Admin - Ver todo" else st.session_state.usuario_actual
    with st.form("f_ind"):
        f_f = st.date_input("Fecha", value=datetime.now())
        f_t = st.selectbox("Tarea", list(COLORES_TAREAS.keys()))
        f_h = st.number_input("Horas", step=0.5); f_n = st.text_input("Nota")
        if st.form_submit_button("Guardar"):
            nueva = {'Fecha': f_f, 'Tarea': f_t, 'Nota': f_n}
            for op in OPERARIOS_FIJOS: nueva[op] = f_h if op == u_c else 0
            st.session_state.cargas = pd.concat([st.session_state.cargas, pd.DataFrame([nueva])], ignore_index=True)
            guardar_df("Cargas", st.session_state.cargas)
            st.success("Guardado"); time.sleep(1); st.rerun()

    st.divider()
    st.subheader(f"Resumen y Detalle para {u_c}")
    df_res = st.session_state.cargas.copy()
    df_res['Fecha'] = pd.to_datetime(df_res['Fecha'], errors='coerce')
    df_res = df_res[df_res[u_c] > 0]
    if not df_res.empty:
        df_res['Mes_N'] = df_res['Fecha'].dt.month; df_res['Año'] = df_res['Fecha'].dt.year; df_res['Mes'] = df_res['Mes_N'].map(MESES_ES)
        cuadro = df_res.groupby(['Año', 'Mes_N', 'Mes', 'Tarea'])[u_c].sum().reset_index()
        st.dataframe(cuadro.sort_values(by=['Año', 'Mes_N'], ascending=False)[['Año', 'Mes', 'Tarea', u_c]], use_container_width=True, hide_index=True)
        
        st.write("**Historial Detallado:**")
        for i, row in df_res.sort_values('Fecha', ascending=False).iterrows():
            c_i, c_d = st.columns([5,1])
            c_i.write(f"**{row['Fecha'].strftime('%d/%m/%Y')}** - {row['Tarea']}: {row[u_c]} hs")
            if c_d.button("Eliminar", key=f"del_{i}"):
                st.session_state.cargas = st.session_state.cargas.drop(i).reset_index(drop=True)
                guardar_df("Cargas", st.session_state.cargas); st.rerun()

elif menu == "Resetear Datos":
    if st.text_input("Escriba BORRAR") == "BORRAR":
        if st.button("Eliminar Todo"):
            guardar_df("Cargas", pd.DataFrame(columns=['Fecha', 'Tarea'] + OPERARIOS_FIJOS + ['Nota']))
            st.rerun()
