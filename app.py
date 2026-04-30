import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import plotly.express as px
import gspread
from google.oauth2.service_account import Credentials
from io import BytesIO
import time

# ===== 1. CONFIGURACIÓN GENERAL =====
st.set_page_config(page_title="CRM Grupo Pressacco", layout="wide")

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

# ===== 2. FUNCIONES PDF (CON COLORES MULTIPLES) =====

def generar_pdf_base(titulo_doc, subtitulo, datos_tablas, incluir_grafico=None):
    from reportlab.lib.pagesizes import letter
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib import colors
    from reportlab.graphics.shapes import Drawing
    from reportlab.graphics.charts.piecharts import Pie
    from reportlab.graphics.charts.legends import Legend
    from reportlab.lib.units import inch

    buf = BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=letter, leftMargin=0.5*inch, rightMargin=0.5*inch)
    s = getSampleStyleSheet()
    
    color_celeste = colors.Color(0, 0.48, 0.73) 
    estilo_titulo = s['Title']
    estilo_titulo.textColor = color_celeste
    estilo_celda = s['Normal']
    estilo_celda.fontSize = 8

    story = [
        Paragraph("GRUPO PRESSACCO", estilo_titulo),
        Paragraph(titulo_doc, s['Heading2']),
        Paragraph(subtitulo, s['Normal']),
        Spacer(1, 15)
    ]

    if incluir_grafico:
        d = Drawing(450, 200)
        pc = Pie()
        pc.x = 50
        pc.y = 25
        pc.width = 130
        pc.height = 130
        
        # Lista de colores vibrantes para el PDF
        lista_colores = [colors.magenta, colors.deepskyblue, colors.lightpink, colors.yellow, 
                         colors.whitesmoke, colors.lightblue, colors.lavender, colors.bisque, 
                         colors.turquoise, colors.lime, colors.hotpink]

        total = sum(incluir_grafico.values())
        pc.data = [round(float(v), 1) for v in incluir_grafico.values()]
        pc.labels = [f"{round((v/total)*100, 1)}%" for v in incluir_grafico.values()]
        pc.sideLabels = 0 
        
        # Asignar colores variados a cada porción
        for i in range(len(pc.data)):
            pc.slices[i].fillColor = lista_colores[i % len(lista_colores)]
            
        leg = Legend()
        leg.x = 220
        leg.y = 150
        leg.alignment = 'right'
        leg.columnMaximum = 12
        leg.fontSize = 7
        # Leyenda con el color correspondiente
        leg.colorNamePairs = [(lista_colores[i % len(lista_colores)], k) for i, k in enumerate(incluir_grafico.keys())]
        
        d.add(pc)
        d.add(leg)
        story.append(d)
        story.append(Spacer(1, 10))

    for titulo_tabla, data in datos_tablas:
        if titulo_tabla:
            story.append(Paragraph(titulo_tabla, s['Heading3']))
        
        data_p = [[Paragraph(str(c), estilo_celda) for c in fila] for fila in data]
        col_w = [2.8*inch] + [1.1*inch] * (len(data[0]) - 1)
        
        t = Table(data_p, colWidths=col_w)
        t.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), color_celeste),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))
        story.append(t)
        story.append(Spacer(1, 15))

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

if 'cargas' not in st.session_state:
    df, _ = cargar_hoja("Cargas")
    st.session_state.cargas = df if not df.empty else pd.DataFrame(columns=['Fecha', 'Tarea'] + OPERARIOS_FIJOS + ['Nota'])

if 'usuario_actual' not in st.session_state:
    st.session_state.usuario_actual = None

if st.session_state.usuario_actual is None:
    st.title("🏛️ CRM Grupo Pressacco")
    u = st.selectbox("Usuario:", ["Seleccionar..."] + OPERARIOS_FIJOS + ["Admin - Ver todo"])
    if st.button("Ingresar") and u != "Seleccionar...":
        st.session_state.usuario_actual = u
        st.rerun()
    st.stop()

# ===== 4. MENÚ Y PANEL =====
opciones = ["📊 Panel de Control", "➕ Cargar Horas", "📁 Carga Masiva", "📜 Protocolo", "⚙️ Reset"] if st.session_state.usuario_actual == "Admin - Ver todo" else ["📊 Panel de Control", "➕ Cargar Mis Horas", "📜 Protocolo"]
menu = st.sidebar.radio("Navegación", opciones)

if st.sidebar.button("Cerrar Sesión"):
    st.session_state.clear(); st.rerun()

if "Panel de Control" in menu:
    st.title("📊 Análisis de Ocupación")
    c1, c2, c3 = st.columns([1,1,2])
    with c1: anio = st.selectbox("Año", [2025, 2026, 2027], index=1)
    with c2: mes = st.selectbox("Mes", list(range(1,13)), format_func=lambda x: MESES_ES[x], index=datetime.now().month-1)
    with c3: p_sel = st.selectbox("Integrante Individual:", OPERARIOS_FIJOS) if st.session_state.usuario_actual == "Admin - Ver todo" else st.session_state.usuario_actual

    df_p = st.session_state.cargas.copy(); df_p['Fecha'] = pd.to_datetime(df_p['Fecha'], errors='coerce')
    
    comp_list = []; hist_pdf = {}
    for i in range(3):
        m_c = mes - i; a_c = anio
        if m_c <= 0: m_c += 12; a_c -= 1
        df_m = df_p[(df_p['Fecha'].dt.month == m_c) & (df_p['Fecha'].dt.year == a_c)]
        total = round(df_m[p_sel].sum(), 1)
        comp_list.append({"Mes": MESES_ES[m_c], "Total": f"{total} hs"})
        hist_pdf[MESES_ES[m_c]] = df_m.groupby('Tarea')[p_sel].sum().to_dict()

    st.subheader(f"📈 Comparativa Trimestral - {p_sel}")
    st.table(pd.DataFrame(comp_list))
    
    if st.button("📥 Descargar Autoevaluación Trimestral (PDF)"):
        tareas_u = sorted(list(set([t for m in hist_pdf for t in hist_pdf[m].keys()])))
        meses_n = list(hist_pdf.keys())
        header = ["Tarea"] + meses_n
        rows = [[t] + [round(float(hist_pdf[m].get(t, 0)), 1) for m in meses_n] for t in tareas_u]
        pdf_t = generar_pdf_base(f"Autoevaluación Trimestral: {p_sel}", "Comparativa de horas por mes", [("Desvío por Tarea", [header] + rows)])
        st.download_button("Guardar Trimestral", pdf_t, f"Trimestral_{p_sel}.pdf")

    st.divider()
    df_act = df_p[(df_p['Fecha'].dt.month == mes) & (df_p['Fecha'].dt.year == anio)]
    res_ind = df_act.groupby('Tarea')[p_sel].sum().round(1).reset_index()
    res_ind = res_ind[res_ind[p_sel]>0]
    
    col_p, col_m = st.columns([2,1])
    with col_p:
        st.plotly_chart(px.pie(res_ind, values=p_sel, names='Tarea', color='Tarea', color_discrete_map=COLORES_TAREAS, title=f"Ocupación Individual {MESES_ES[mes]}"), use_container_width=True)
    with col_m:
        if st.button("📥 Reporte Mensual Individual (PDF)"):
            dict_pie = res_ind.set_index('Tarea')[p_sel].to_dict()
            datos_t = [["Tarea", "Horas"]] + [[r['Tarea'], round(r[p_sel], 1)] for _, r in res_ind.iterrows()]
            pdf_m = generar_pdf_base(f"Reporte Mensual: {p_sel}", f"{MESES_ES[mes]} {anio}", [("Detalle", datos_t)], incluir_grafico=dict_pie)
            st.download_button("Guardar Mensual", pdf_m, f"Mensual_{p_sel}.pdf")

    if st.session_state.usuario_actual == "Admin - Ver todo":
        st.divider()
        st.subheader("🌐 Visión Global del Estudio")
        df_eq = df_act.melt(id_vars=['Fecha', 'Tarea'], value_vars=OPERARIOS_FIJOS, var_name='Op', value_name='Hs')
        res_eq = df_eq.groupby('Tarea')['Hs'].sum().reset_index()
        st.plotly_chart(px.pie(res_eq, values='Hs', names='Tarea', color='Tarea', color_discrete_map=COLORES_TAREAS, title="Total Horas Estudio"), use_container_width=True)
        if st.button("📥 Descargar Reporte Global (PDF)"):
            dict_g = res_eq.set_index('Tarea')['Hs'].to_dict()
            datos_g = [["Tarea", "Suma Horas"]] + [[t, round(h, 1)] for t, h in dict_g.items()]
            pdf_g = generar_pdf_base("Reporte Global de Equipo", f"Estudio Completo - {MESES_ES[mes]}", [("Totales", datos_g)], incluir_grafico=dict_g)
            st.download_button("Guardar Global", pdf_g, "Reporte_Global.pdf")

elif "Cargar" in menu:
    st.title("➕ Registro de Horas")
    u_c = st.session_state.usuario_actual if st.session_state.usuario_actual != "Admin - Ver todo" else st.selectbox("Persona:", OPERARIOS_FIJOS)
    with st.form("f_ind"):
        f_f = st.date_input("Fecha"); f_t = st.selectbox("Tarea", list(COLORES_TAREAS.keys())); f_h = st.number_input("Horas", step=0.5); f_n = st.text_input("Nota")
        if st.form_submit_button("Guardar"):
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
