import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import plotly.express as px
import plotly.graph_objects as go
import gspread
from google.oauth2.service_account import Credentials
from io import BytesIO
import time

# ===== 1. CONFIGURACIÓN ESTÉTICA Y GENERAL =====
st.set_page_config(page_title="CRM Grupo Pressacco", layout="wide")

# CSS personalizado para mejorar la interfaz
st.markdown("""
    <style>
    .main { background-color: #f5f7f9; }
    .stMetric { background-color: #ffffff; padding: 15px; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    .stButton>button { width: 100%; border-radius: 5px; height: 3em; background-color: #007bff; color: white; }
    .sidebar .sidebar-content { background-image: linear-gradient(#2e7bcf,#2e7bcf); color: white; }
    </style>
    """, unsafe_allow_html=True)

COLORES_TAREAS = {
    "DOCUMENTACIÓN CARGA": "#FFB6C1", "DOCUMENTACIÓN CONTROL": "#FF69B4",
    "IMPUESTOS": "#FF00FF", "SUELDOS": "#FFFF00", "CONTABILIDAD": "#00FF00",
    "ATENCION AL CLIENTE": "#00BFFF", "TAREAS NO RUTINARIAS": "#ADD8E6",
    "DISPONIBLE": "#FFFFFF", "REUNIONES DE EQUIPO": "#E6E6FA",
    "INASISTENCIA POR EXAMEN O TRAMITE": "#FFDAB9",
    "PLANIFICACIONES/ORGANIZACIÓN/PROCEDIMIENTO S/INFORMES": "#40E0D0"
}

OPERARIOS_FIJOS = ["Natalia", "Maximiliano", "Athina", "Johana"]
HORAS_DIA_LABORAL = 6
TAREAS_DISPONIBLE_TIPO = ["DISPONIBLE", "PLANIFICACIONES/ORGANIZACIÓN/PROCEDIMIENTO S/INFORMES", "INASISTENCIA POR EXAMEN O TRAMITE"]
MESES_ES = {1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
            7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"}

# ===== 2. FUNCIONES PDF (REFORZADAS) =====
def generar_pdf_protocolo_total():
    from reportlab.lib.pagesizes import letter
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, ListFlowable, ListItem
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib.enums import TA_CENTER
    buf = BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=letter)
    s = getSampleStyleSheet()
    story = [
        Paragraph("GRUPO PRESSACCO - Protocolo Integral", s['Title']),
        Paragraph("<i>Elegimos sumar</i>", s['Normal']),
        Spacer(1, 20),
        Paragraph("1. Acceso y Privacidad", s['Heading2']),
        Paragraph("Es obligatorio seleccionar su propio nombre al ingresar. Nunca cargue horas en el panel de un compañero para evitar desvíos en la capacidad real del equipo.", s['Normal']),
        Spacer(1, 10),
        Paragraph("2. Reglas de Carga Diaria", s['Heading2']),
        ListFlowable([
            ListItem(Paragraph("Carga obligatoria antes de las 15:00 hs de cada día laboral.", s['Normal'])),
            ListItem(Paragraph("La suma total debe ser siempre de 6 horas.", s['Normal'])),
            ListItem(Paragraph("Si no hay tareas operativas, use 'DISPONIBLE'.", s['Normal'])),
        ], bulletType='bullet'),
        Spacer(1, 10),
        Paragraph("3. Análisis de Desvíos", s['Heading2']),
        Paragraph("Utilice el Panel de Control para comparar su eficiencia entre meses. Si su disponibilidad baja del 10% (Rojo), informe de inmediato.", s['Normal']),
    ]
    doc.build(story)
    buf.seek(0)
    return buf

def generar_pdf_trimestral(nombre, comp_data, historial_tareas):
    from reportlab.lib.pagesizes import letter
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib import colors
    buf = BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=letter)
    s = getSampleStyleSheet()
    story = [Paragraph("GRUPO PRESSACCO", s['Title']), Paragraph(f"Autoevaluación Trimestral: {nombre}", s['Heading1']), Spacer(1, 15)]
    data_gen = [["Mes", "Total", "Eficiencia", "Disponible"]]
    for item in comp_data: data_gen.append([item['Mes'], item['Total'], item['Eficiencia'], item['Disponible']])
    tg = Table(data_gen, colWidths=[120, 100, 100, 100])
    tg.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.grey), ('TEXTCOLOR',(0,0),(-1,0),colors.whitesmoke), ('GRID',(0,0),(-1,-1),1,colors.black)]))
    story.append(tg); story.append(Spacer(1, 20))
    story.append(Paragraph("Evolución de Horas por Tarea", s['Heading2']))
    tareas_u = sorted(list(set([t for m in historial_tareas for t in historial_tareas[m].keys()])))
    meses_n = list(historial_tareas.keys())
    data_t = [["Tarea"] + meses_n]
    for t in tareas_u:
        fila = [t]
        for m in meses_n: fila.append(f"{historial_tareas[m].get(t, 0):.1f} hs")
        data_t.append(fila)
    tt = Table(data_t, colWidths=[200, 80, 80, 80])
    tt.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.lightgrey), ('GRID',(0,0),(-1,-1),1,colors.black), ('FONTSIZE',(0,0),(-1,-1), 8)]))
    story.append(tt)
    doc.build(story); buf.seek(0); return buf

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
        ws.clear(); df_c = df.copy()
        for c in df_c.columns:
            if 'fecha' in c.lower(): df_c[c] = pd.to_datetime(df_c[c], errors='coerce').dt.strftime('%d/%m/%Y')
        ws.update([df_c.columns.values.tolist()] + df_c.fillna('').astype(str).values.tolist())
        st.cache_data.clear(); return True
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
    st.title("🏛️ CRM Grupo Pressacco")
    u = st.selectbox("Identifíquese para continuar:", ["Seleccionar..."] + OPERARIOS_FIJOS + ["Admin - Ver todo"])
    if st.button("Ingresar al Sistema") and u != "Seleccionar...":
        st.session_state.usuario_actual = u
        st.rerun()
    st.stop()

# ===== 5. MENÚ SIDEBAR =====
st.sidebar.title(f"👤 {st.session_state.usuario_actual}")
opciones = ["📊 Panel de Control", "➕ Cargar Horas", "📁 Carga Masiva", "📜 Protocolo", "⚙️ Reset"] if st.session_state.usuario_actual == "Admin - Ver todo" else ["📊 Panel de Control", "➕ Cargar Mis Horas", "📜 Protocolo"]
menu = st.sidebar.radio("Navegación", opciones)
if st.sidebar.button("Cerrar Sesión"):
    st.session_state.clear(); st.rerun()

# ===== 6. SECCIONES =====
if "Panel de Control" in menu:
    st.title("📊 Análisis de Capacidad")
    with st.container():
        c1, c2, c3 = st.columns([1,1,2])
        with c1: anio = st.selectbox("Año", [2025, 2026, 2027], index=1)
        with c2: mes = st.selectbox("Mes", list(range(1,13)), format_func=lambda x: MESES_ES[x], index=datetime.now().month-1)
        with c3: p_sel = st.selectbox("Integrante:", OPERARIOS_FIJOS) if st.session_state.usuario_actual == "Admin - Ver todo" else st.session_state.usuario_actual

    df_p = st.session_state.cargas.copy(); df_p['Fecha'] = pd.to_datetime(df_p['Fecha'], errors='coerce')
    
    st.subheader(f"📈 Comparativa Trimestral - {p_sel}")
    comp_list = []; hist_tareas_pdf = {}
    for i in range(3):
        m_c = mes - i; a_c = anio
        if m_c <= 0: m_c += 12; a_c -= 1
        ini = datetime(a_c, m_c, 1).date(); fin = (datetime(a_c, m_c+1, 1) if m_c < 12 else datetime(a_c+1, 1, 1)).date() - timedelta(days=1)
        cap = len(pd.bdate_range(start=ini, end=fin, freq='C', holidays=FERIADOS)) * HORAS_DIA_LABORAL
        df_m = df_p[(df_p['Fecha'].dt.month == m_c) & (df_p['Fecha'].dt.year == a_c)]
        total = round(df_m[p_sel].sum(), 1); h_prod = df_m[~df_m['Tarea'].isin(TAREAS_DISPONIBLE_TIPO)][p_sel].sum()
        h_disp = df_m[df_m['Tarea'].isin(TAREAS_DISPONIBLE_TIPO)][p_sel].sum()
        comp_list.append({"Mes": MESES_ES[m_c], "Total": f"{total} hs", "Eficiencia": f"{(h_prod/cap*100 if cap>0 else 0):.1f}%", "Disponible": f"{(h_disp/total*100 if total>0 else 0):.1f}%"})
        hist_tareas_pdf[MESES_ES[m_c]] = df_m.groupby('Tarea')[p_sel].sum().to_dict()
    
    st.table(pd.DataFrame(comp_list))
    st.download_button("📥 Descargar Autoevaluación Trimestral (PDF)", data=generar_pdf_trimestral(p_sel, comp_list, hist_tareas_pdf), file_name=f"Trimestral_{p_sel}.pdf")

    st.divider()
    df_act = df_p[(df_p['Fecha'].dt.month == mes) & (df_p['Fecha'].dt.year == anio)]
    res_t = df_act.groupby('Tarea')[p_sel].sum().round(1).reset_index(); res_t.columns = ['Tarea', 'Hs']
    c_g, c_m = st.columns([2,1])
    with c_g: st.plotly_chart(px.pie(res_t[res_t['Hs']>0], values='Hs', names='Tarea', color='Tarea', color_discrete_map=COLORES_TAREAS, title=f"Distribución {MESES_ES[mes]}"), use_container_width=True)
    with c_m:
        st.metric("Eficiencia Mes", comp_list[0]["Eficiencia"])
        st.metric("Disponible Mes", comp_list[0]["Disponible"])

    if st.session_state.usuario_actual == "Admin - Ver todo":
        st.divider(); st.subheader("🌐 Visión Global del Estudio")
        df_eq = df_act.melt(id_vars=['Fecha', 'Tarea'], value_vars=OPERARIOS_FIJOS, var_name='Op', value_name='Hs')
        st.plotly_chart(px.pie(df_eq.groupby('Tarea')['Hs'].sum().reset_index(), values='Hs', names='Tarea', color='Tarea', color_discrete_map=COLORES_TAREAS, title="Total Horas Estudio"), use_container_width=True)

elif "Cargar" in menu:
    st.title("➕ Registro de Actividad")
    u_c = st.session_state.usuario_actual if st.session_state.usuario_actual != "Admin - Ver todo" else st.selectbox("Persona:", OPERARIOS_FIJOS)
    with st.form("f_ind"):
        f_f = st.date_input("Fecha"); f_t = st.selectbox("Tarea", list(COLORES_TAREAS.keys())); f_h = st.number_input("Horas", step=0.5); f_n = st.text_input("Nota")
        if st.form_submit_button("Guardar Registro"):
            nueva = {'Fecha': f_f, 'Tarea': f_t, 'Nota': f_n}
            for op in OPERARIOS_FIJOS: nueva[op] = f_h if op == u_c else 0
            st.session_state.cargas = pd.concat([st.session_state.cargas, pd.DataFrame([nueva])], ignore_index=True)
            guardar_df("Cargas", st.session_state.cargas); st.success("¡Guardado!"); time.sleep(1); st.rerun()

    st.divider(); df_r = st.session_state.cargas.copy(); df_r['Fecha'] = pd.to_datetime(df_r['Fecha'], errors='coerce')
    df_r = df_r[df_r[u_c] > 0]
    if not df_r.empty:
        df_r['Mes_N'] = df_r['Fecha'].dt.month; df_r['Año'] = df_r['Fecha'].dt.year; df_r['Mes'] = df_r['Mes_N'].map(MESES_ES)
        st.dataframe(df_r.groupby(['Año', 'Mes_N', 'Mes', 'Tarea'])[u_c].sum().round(1).reset_index().sort_values(by=['Año', 'Mes_N'], ascending=False)[['Año', 'Mes', 'Tarea', u_c]], use_container_width=True, hide_index=True)
        for i, row in df_r.sort_values('Fecha', ascending=False).iterrows():
            c_i, c_d = st.columns([5,1]); c_i.write(f"**{row['Fecha'].strftime('%d/%m/%Y')}** - {row['Tarea']}: {row[u_c]:.1f} hs")
            if c_d.button("Eliminar", key=f"del_{i}"):
                st.session_state.cargas = st.session_state.cargas.drop(i).reset_index(drop=True)
                guardar_df("Cargas", st.session_state.cargas); st.rerun()

elif "Carga Masiva" in menu:
    st.title("📁 Reparto de Horas (Admin)")
    with st.form("f_masiva"):
        u_m = st.selectbox("Operario", OPERARIOS_FIJOS); t_m = st.selectbox("Tarea", list(COLORES_TAREAS.keys()))
        f_i = st.date_input("Desde"); f_f = st.date_input("Hasta"); h_t = st.number_input("Horas Totales", min_value=0.0)
        if st.form_submit_button("Distribuir Horas"):
            dias = pd.bdate_range(start=f_i, end=f_f, freq='C', holidays=FERIADOS)
            if len(dias) > 0:
                h_d = round(h_t / len(dias), 2); filas = []
                for d in dias:
                    f = {'Fecha': d, 'Tarea': t_m, 'Nota': 'Carga Masiva'}
                    for o in OPERARIOS_FIJOS: f[o] = h_d if o == u_m else 0
                    filas.append(f)
                st.session_state.cargas = pd.concat([st.session_state.cargas, pd.DataFrame(filas)], ignore_index=True)
                guardar_df("Cargas", st.session_state.cargas); st.success("Carga masiva lista"); time.sleep(1); st.rerun()

elif "Protocolo" in menu:
    st.title("📜 Protocolo: Grupo Pressacco")
    st.info("Elegimos sumar")
    st.markdown("""
    ### 1. Acceso Personalizado
    Cada integrante debe ingresar con su propio usuario. No comparta sesiones.
    ### 2. Carga de Horas
    - **Diario:** Cargar antes de las 15:00 hs.
    - **Total:** Debe sumar 6 horas cada día.
    ### 3. Tareas
    - Si no tiene tareas asignadas, use **DISPONIBLE**.
    - Sea preciso con las tareas operativas (Impuestos, Sueldos, etc.).
    """)
    st.download_button("📥 Descargar Protocolo Integral (PDF)", data=generar_pdf_protocolo_total(), file_name="Protocolo_Pressacco.pdf")

elif "Reset" in menu:
    if st.text_input("Escriba BORRAR") == "BORRAR":
        if st.button("Eliminar Todo"): guardar_df("Cargas", pd.DataFrame(columns=['Fecha', 'Tarea'] + OPERARIOS_FIJOS + ['Nota'])); st.rerun()
