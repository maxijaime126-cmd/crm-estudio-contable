import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import plotly.express as px
import plotly.graph_objects as go
import gspread
from google.oauth2.service_account import Credentials
from io import BytesIO
import time

# ===== 1. CONFIGURACIÓN GENERAL =====
st.set_page_config(page_title="CRM Grupo Pressacco", layout="wide")

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

# ===== 2. FUNCIONES PDF (REPORTE MENSUAL Y TRIMESTRAL) =====

def generar_pdf_trimestral(nombre, comp_data, historial_tareas):
    from reportlab.lib.pagesizes import letter
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib import colors
    
    buf = BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=letter)
    s = getSampleStyleSheet()
    
    story = [
        Paragraph("GRUPO PRESSACCO", s['Title']),
        Paragraph(f"Autoevaluación Trimestral: {nombre}", s['Heading1']),
        Spacer(1, 15),
    ]
    
    # Tabla Comparativa General
    story.append(Paragraph("Resumen de Métricas Principales", s['Heading2']))
    data_gen = [["Mes", "Carga Total", "Eficiencia", "Disponible"]]
    for item in comp_data:
        data_gen.append([item['Mes'], item['Total'], item['Eficiencia'], item['Disponible']])
    
    tg = Table(data_gen, colWidths=[120, 100, 100, 100])
    tg.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.grey), ('TEXTCOLOR',(0,0),(-1,0),colors.whitesmoke), ('GRID',(0,0),(-1,-1),1,colors.black)]))
    story.append(tg)
    story.append(Spacer(1, 20))
    
    # Tabla Comparativa de Tareas (Desvío de horas)
    story.append(Paragraph("Evolución de Horas por Tarea", s['Heading2']))
    tareas_unicas = sorted(list(set([t for m in historial_tareas for t in historial_tareas[m].keys()])))
    meses_nombres = list(historial_tareas.keys())
    
    data_tareas = [["Tarea"] + meses_nombres]
    for tarea in tareas_unicas:
        fila = [tarea]
        for m in meses_nombres:
            fila.append(f"{historial_tareas[m].get(tarea, 0):.1f} hs")
        data_tareas.append(fila)
        
    tt = Table(data_tareas, colWidths=[200, 80, 80, 80])
    tt.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.lightgrey), ('GRID',(0,0),(-1,-1),1,colors.black), ('FONTSIZE',(0,0),(-1,-1), 8)]))
    story.append(tt)
    
    doc.build(story)
    buf.seek(0)
    return buf

def generar_pdf_reporte_completo(nombre, mes, anio, total_hs, eficiencia, estado, df_resumen):
    from reportlab.lib.pagesizes import letter
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib import colors
    buf = BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=letter)
    s = getSampleStyleSheet()
    story = [
        Paragraph("GRUPO PRESSACCO", s['Title']),
        Paragraph(f"Reporte Mensual: {nombre}", s['Heading1']),
        Paragraph(f"Período: {MESES_ES[mes]} {anio}", s['Normal']),
        Spacer(1, 15),
    ]
    data_m = [["Métrica", "Valor"], ["Total Horas", f"{total_hs:.1f} hs"], ["Eficiencia", f"{eficiencia:.1f}%"], ["Estado", estado]]
    tm = Table(data_m, colWidths=[200, 200])
    tm.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.grey), ('TEXTCOLOR',(0,0),(-1,0),colors.whitesmoke), ('GRID',(0,0),(-1,-1),1,colors.black)]))
    story.append(tm)
    story.append(Spacer(1, 20))
    data_t = [["Tarea", "Horas", "%"]]
    for _, row in df_resumen.iterrows():
        porc = (row['Hs'] / total_hs * 100) if total_hs > 0 else 0
        data_t.append([row['Tarea'], f"{row['Hs']:.1f} hs", f"{porc:.1f}%"])
    tt = Table(data_t, colWidths=[250, 80, 80])
    tt.setStyle(TableStyle([('GRID',(0,0),(-1,-1),1,colors.black)]))
    story.append(tt)
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
    st.title("CRM Grupo Pressacco")
    u = st.selectbox("Usuario:", ["Seleccionar..."] + OPERARIOS_FIJOS + ["Admin - Ver todo"])
    if st.button("Ingresar") and u != "Seleccionar...":
        st.session_state.usuario_actual = u
        st.rerun()
    st.stop()

# ===== 5. MENÚ SIDEBAR =====

opciones = ["Panel de Control", "Cargar Horas", "Carga Masiva", "Protocolo de Trabajo", "Resetear Datos"] if st.session_state.usuario_actual == "Admin - Ver todo" else ["Panel de Control", "Cargar Mis Horas", "Protocolo de Trabajo"]
menu = st.sidebar.radio("Menú", opciones)

if st.sidebar.button("Cerrar Sesión"):
    st.session_state.clear()
    st.rerun()

# ===== 6. PANEL DE CONTROL =====

if menu == "Panel de Control":
    st.title("Análisis y Autoevaluación")
    c1, c2 = st.columns(2)
    with c1: anio = st.selectbox("Año", [2025, 2026, 2027], index=1)
    with c2: mes = st.selectbox("Mes", list(range(1,13)), format_func=lambda x: MESES_ES[x], index=datetime.now().month-1)

    p_sel = st.selectbox("Integrante Individual:", OPERARIOS_FIJOS) if st.session_state.usuario_actual == "Admin - Ver todo" else st.session_state.usuario_actual
    df_p = st.session_state.cargas.copy()
    df_p['Fecha'] = pd.to_datetime(df_p['Fecha'], errors='coerce')

    # Cuadro Comparativo Trimestral y Captura para PDF
    st.subheader(f"Comparativa Trimestral - {p_sel}")
    comp_list = []
    historial_tareas_pdf = {} # Guardamos para el PDF
    
    for i in range(3):
        m_c = mes - i
        a_c = anio
        if m_c <= 0: m_c += 12; a_c -= 1
        
        ini = datetime(a_c, m_c, 1).date()
        fin = (datetime(a_c, m_c+1, 1) if m_c < 12 else datetime(a_c+1, 1, 1)).date() - timedelta(days=1)
        cap = len(pd.bdate_range(start=ini, end=fin, freq='C', holidays=FERIADOS)) * HORAS_DIA_LABORAL
        
        df_m = df_p[(df_p['Fecha'].dt.month == m_c) & (df_p['Fecha'].dt.year == a_c)]
        total = round(df_m[p_sel].sum(), 1)
        h_prod = df_m[~df_m['Tarea'].isin(TAREAS_DISPONIBLE_TIPO)][p_sel].sum()
        h_disp = df_m[df_m['Tarea'].isin(TAREAS_DISPONIBLE_TIPO)][p_sel].sum()
        efic = (h_prod / cap * 100) if cap > 0 else 0
        dispon = (h_disp / total * 100) if total > 0 else 0
        
        nombre_mes = MESES_ES[m_c]
        comp_list.append({"Mes": nombre_mes, "Total": f"{total} hs", "Eficiencia": f"{efic:.1f}%", "Disponible": f"{dispon:.1f}%"})
        
        # Detalle de tareas para el PDF
        historial_tareas_pdf[nombre_mes] = df_m.groupby('Tarea')[p_sel].sum().to_dict()
    
    st.table(pd.DataFrame(comp_list))
    
    if st.button("📥 Generar Autoevaluación Trimestral (PDF)"):
        pdf_tri = generar_pdf_trimestral(p_sel, comp_list, historial_tareas_pdf)
        st.download_button("Descargar PDF Trimestral", pdf_tri, f"Autoevaluacion_Trimestral_{p_sel}.pdf")

    st.divider()

    # Detalle Individual del Mes Actual
    df_actual = df_p[(df_p['Fecha'].dt.month == mes) & (df_p['Fecha'].dt.year == anio)]
    res_t = df_actual.groupby('Tarea')[p_sel].sum().round(1).reset_index()
    res_t.columns = ['Tarea', 'Hs']
    res_t = res_t[res_t['Hs'] > 0]
    
    col_g, col_m = st.columns([2,1])
    with col_g:
        st.plotly_chart(px.pie(res_t, values='Hs', names='Tarea', color='Tarea', color_discrete_map=COLORES_TAREAS, title=f"Ocupación {p_sel} ({MESES_ES[mes]})"), use_container_width=True)
    with col_m:
        st.metric("Eficiencia Mes", comp_list[0]["Eficiencia"])
        st.metric("Disponible Mes", comp_list[0]["Disponible"])
        if st.button("Generar Reporte Mensual PDF"):
            pdf_mensual = generar_pdf_reporte_completo(p_sel, mes, anio, df_actual[p_sel].sum(), float(comp_list[0]["Eficiencia"].replace('%','')), "Analizado", res_t)
            st.download_button("Descargar Reporte Mensual", pdf_mensual, f"Reporte_{p_sel}_{MESES_ES[mes]}.pdf")

    # --- EQUIPO TOTAL (SOLO ADMIN) ---
    if st.session_state.usuario_actual == "Admin - Ver todo":
        st.divider()
        st.subheader("Visión General del Equipo (Estudio Completo)")
        df_eq = df_actual.melt(id_vars=['Fecha', 'Tarea'], value_vars=OPERARIOS_FIJOS, var_name='Op', value_name='Hs')
        res_eq = df_eq.groupby('Tarea')['Hs'].sum().round(1).reset_index()
        
        c_eq1, c_eq2 = st.columns([2, 1])
        with c_eq1:
            st.plotly_chart(px.pie(res_eq, values='Hs', names='Tarea', color='Tarea', color_discrete_map=COLORES_TAREAS, title="Distribución Total del Estudio"), use_container_width=True)
        with c_eq2:
            st.write("**Total por Integrante:**")
            res_ops = []
            for o in OPERARIOS_FIJOS:
                res_ops.append({'Operario': o, 'Horas': f"{df_actual[o].sum().round(1)} hs"})
            st.table(pd.DataFrame(res_ops))

# ===== 7. CARGA DE HORAS =====

elif "Cargar" in menu:
    st.title("Registro de Horas")
    u_c = st.session_state.usuario_actual if st.session_state.usuario_actual != "Admin - Ver todo" else st.selectbox("Persona:", OPERARIOS_FIJOS)
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
    df_r = st.session_state.cargas.copy()
    df_r['Fecha'] = pd.to_datetime(df_r['Fecha'], errors='coerce')
    df_r = df_r[df_r[u_c] > 0]
    if not df_r.empty:
        df_r['Mes_N'] = df_r['Fecha'].dt.month; df_r['Año'] = df_r['Fecha'].dt.year; df_r['Mes'] = df_r['Mes_N'].map(MESES_ES)
        st.dataframe(df_r.groupby(['Año', 'Mes_N', 'Mes', 'Tarea'])[u_c].sum().round(1).reset_index().sort_values(by=['Año', 'Mes_N'], ascending=False)[['Año', 'Mes', 'Tarea', u_c]], use_container_width=True, hide_index=True)
        for i, row in df_r.sort_values('Fecha', ascending=False).iterrows():
            c_i, c_d = st.columns([5,1])
            c_i.write(f"**{row['Fecha'].strftime('%d/%m/%Y')}** - {row['Tarea']}: {row[u_c]:.1f} hs")
            if c_d.button("Eliminar", key=f"del_{i}"):
                st.session_state.cargas = st.session_state.cargas.drop(i).reset_index(drop=True)
                guardar_df("Cargas", st.session_state.cargas); st.rerun()

# ===== 8. CARGA MASIVA =====

elif menu == "Carga Masiva":
    st.title("Carga Masiva (Admin)")
    with st.form("f_masiva"):
        u_m = st.selectbox("Operario", OPERARIOS_FIJOS); t_m = st.selectbox("Tarea", list(COLORES_TAREAS.keys()))
        f_i = st.date_input("Inicio"); f_f = st.date_input("Fin"); h_t = st.number_input("Horas Totales", min_value=0.0)
        if st.form_submit_button("Repartir Horas"):
            dias = pd.bdate_range(start=f_i, end=f_f, freq='C', holidays=FERIADOS)
            if len(dias) > 0:
                h_d = round(h_t / len(dias), 2)
                filas = []
                for d in dias:
                    f = {'Fecha': d, 'Tarea': t_m, 'Nota': 'Carga Masiva'}
                    for o in OPERARIOS_FIJOS: f[o] = h_d if o == u_m else 0
                    filas.append(f)
                st.session_state.cargas = pd.concat([st.session_state.cargas, pd.DataFrame(filas)], ignore_index=True)
                guardar_df("Cargas", st.session_state.cargas); st.success("Carga masiva lista"); time.sleep(1); st.rerun()

elif menu == "Protocolo de Trabajo":
    st.title("📖 Protocolo: Grupo Pressacco")
    st.info("Elegimos sumar")
    from reportlab.lib.pagesizes import letter
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, ListFlowable, ListItem
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib.enums import TA_CENTER
    buf = BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=letter)
    s = getSampleStyleSheet()
    style_lema = getSampleStyleSheet()['Normal']
    style_lema.alignment = TA_CENTER
    style_lema.fontName = 'Helvetica-Oblique'
    story = [
        Paragraph("GRUPO PRESSACCO", s['Title']),
        Paragraph("<i>Elegimos sumar</i>", style_lema),
        Spacer(1, 20),
        Paragraph("Protocolo de Trabajo del Sistema", s['Heading1']),
        Spacer(1, 12),
        Paragraph("Cada integrante debe cargar sus horas diariamente antes de las 15:00 hs. El total debe sumar 6 horas.", s['Normal']),
    ]
    doc.build(story)
    buf.seek(0)
    st.download_button("Descargar Protocolo PDF", data=buf, file_name="Protocolo_Pressacco.pdf")

elif menu == "Resetear Datos":
    if st.text_input("Escriba BORRAR") == "BORRAR":
        if st.button("Eliminar Todo"):
            guardar_df("Cargas", pd.DataFrame(columns=['Fecha', 'Tarea'] + OPERARIOS_FIJOS + ['Nota'])); st.rerun()
