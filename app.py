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

# ===== 2. FUNCIONES PDF =====

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
    estilo_negrita = s['Normal']
    estilo_negrita.fontSize = 8
    estilo_negrita.fontName = 'Helvetica-Bold'

    story = [
        Paragraph("GRUPO PRESSACCO", estilo_titulo),
        Paragraph(titulo_doc, s['Heading2']),
        Paragraph(subtitulo, s['Normal']),
        Spacer(1, 15)
    ]

    if incluir_grafico:
        d = Drawing(450, 200)
        pc = Pie()
        pc.x = 50; pc.y = 25; pc.width = 130; pc.height = 130
        lista_colores = [colors.magenta, colors.deepskyblue, colors.lightpink, colors.yellow, 
                         colors.whitesmoke, colors.lightblue, colors.lavender, colors.bisque, 
                         colors.turquoise, colors.lime, colors.hotpink]
        total_h = sum(incluir_grafico.values())
        pc.data = [round(float(v), 1) for v in incluir_grafico.values()]
        pc.labels = [f"{round((v/total_h)*100, 1)}%" if total_h > 0 else "0%" for v in incluir_grafico.values()]
        for i in range(len(pc.data)):
            pc.slices[i].fillColor = lista_colores[i % len(lista_colores)]
        leg = Legend()
        leg.x = 220; leg.y = 150; leg.alignment = 'right'; leg.columnMaximum = 12; leg.fontSize = 7
        leg.colorNamePairs = [(lista_colores[i % len(lista_colores)], k) for i, k in enumerate(incluir_grafico.keys())]
        d.add(pc); d.add(leg); story.append(d); story.append(Spacer(1, 10))

    for titulo_tabla, data in datos_tablas:
        if titulo_tabla:
            story.append(Paragraph(titulo_tabla, s['Heading3']))
        data_p = []
        for fila in data:
            estilo = estilo_negrita if str(fila[0]).upper() == "TOTAL" else estilo_celda
            data_p.append([Paragraph(str(c), estilo) for c in fila])
        col_w = [2.5*inch] + [0.8*inch]*(len(data[0])-1)
        t = Table(data_p, colWidths=col_w)
        estilo_t = [
            ('BACKGROUND', (0, 0), (-1, 0), color_celeste),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]
        if str(data[-1][0]).upper() == "TOTAL":
            estilo_t.append(('BACKGROUND', (0, -1), (-1, -1), colors.lightgrey))
        t.setStyle(TableStyle(estilo_t))
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

# ===== 4. LÓGICA DE ALERTA DE CARGA =====
def mostrar_alerta_faltante(usuario):
    if usuario == "Admin - Ver todo": return
    hoy = datetime.now()
    inicio_mes = datetime(hoy.year, hoy.month, 1).date()
    fin_mes = hoy.date()
    # Calcular días hábiles hasta hoy
    dias_habiles = len(pd.bdate_range(start=inicio_mes, end=fin_mes, freq='C', holidays=FERIADOS))
    horas_objetivo = dias_habiles * HORAS_DIA_LABORAL
    
    df_u = st.session_state.cargas.copy()
    df_u['Fecha'] = pd.to_datetime(df_u['Fecha'], errors='coerce')
    horas_cargadas = df_u[(df_u['Fecha'].dt.month == hoy.month) & (df_u['Fecha'].dt.year == hoy.year)][usuario].sum()
    
    if horas_cargadas < horas_objetivo:
        faltan = round(horas_objetivo - horas_cargadas, 1)
        st.warning(f"⚠️ **Aviso de Carga:** Hola {usuario}, te faltan cargar **{faltan} horas** para completar lo correspondiente al mes de {MESES_ES[hoy.month]} hasta hoy.")

# ===== 5. LOGIN =====
if st.session_state.usuario_actual is None:
    st.title("🏛️ CRM Grupo Pressacco")
    u = st.selectbox("Usuario:", ["Seleccionar..."] + OPERARIOS_FIJOS + ["Admin - Ver todo"])
    if st.button("Ingresar") and u != "Seleccionar...":
        st.session_state.usuario_actual = u
        st.rerun()
    st.stop()

# ===== 6. NAVEGACIÓN =====
opciones = ["📊 Panel de Control", "➕ Cargar Horas", "📁 Carga Masiva", "📜 Protocolo", "⚙️ Reset"] if st.session_state.usuario_actual == "Admin - Ver todo" else ["📊 Panel de Control", "➕ Cargar Mis Horas", "📜 Protocolo"]
menu = st.sidebar.radio("Navegación", opciones)

if st.sidebar.button("Cerrar Sesión"):
    st.session_state.clear(); st.rerun()

# Mostrar alerta al inicio si no es admin
mostrar_alerta_faltante(st.session_state.usuario_actual)

# ===== 7. PANEL DE CONTROL =====
if "Panel de Control" in menu:
    st.title("📊 Análisis de Ocupación")
    c1, c2, c3 = st.columns([1,1,2])
    with c1: anio = st.selectbox("Año", [2025, 2026, 2027], index=1)
    with c2: mes = st.selectbox("Mes", list(range(1,13)), format_func=lambda x: MESES_ES[x], index=datetime.now().month-1)
    with c3: p_sel = st.selectbox("Integrante:", OPERARIOS_FIJOS) if st.session_state.usuario_actual == "Admin - Ver todo" else st.session_state.usuario_actual

    df_p = st.session_state.cargas.copy(); df_p['Fecha'] = pd.to_datetime(df_p['Fecha'], errors='coerce')
    
    st.subheader(f"📈 Comparativa Trimestral - {p_sel}")
    comp_list = []; hist_pdf = {}
    for i in range(3):
        m_c = mes - i; a_c = anio
        if m_c <= 0: m_c += 12; a_c -= 1
        df_m = df_p[(df_p['Fecha'].dt.month == m_c) & (df_p['Fecha'].dt.year == a_c)]
        total = round(df_m[p_sel].sum(), 1)
        comp_list.append({"Mes": MESES_ES[m_c], "Total": f"{total} hs"})
        hist_pdf[MESES_ES[m_c]] = df_m.groupby('Tarea')[p_sel].sum().to_dict()

    st.table(pd.DataFrame(comp_list))
    
    if st.button("📥 Descargar Autoevaluación Trimestral (PDF)"):
        tareas_u = sorted(list(set([t for m in hist_pdf for t in hist_pdf[m].keys()])))
        meses_n = list(hist_pdf.keys())
        header = ["Tarea"] + meses_n
        rows = []; totales_m = [0.0] * len(meses_n)
        for t in tareas_u:
            fila = [t]
            for idx, m in enumerate(meses_n):
                val = round(float(hist_pdf[m].get(t, 0)), 1)
                fila.append(val); totales_m[idx] += val
            rows.append(fila)
        rows.append(["TOTAL"] + [round(x, 1) for x in totales_m])
        pdf_t = generar_pdf_base(f"Autoevaluación Trimestral: {p_sel}", "Comparativa mensual", [("Desvío por Tarea", [header] + rows)])
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
            tot_h = res_ind[p_sel].sum()
            dict_pie = res_ind.set_index('Tarea')[p_sel].to_dict()
            datos_t = [["Tarea", "Horas", "%"]]
            for _, r in res_ind.iterrows():
                porc = f"{round((r[p_sel]/tot_h)*100, 1)}%" if tot_h > 0 else "0%"
                datos_t.append([r['Tarea'], round(r[p_sel], 1), porc])
            datos_t.append(["TOTAL", round(tot_h, 1), "100%"])
            pdf_m = generar_pdf_base(f"Reporte Mensual: {p_sel}", f"{MESES_ES[mes]} {anio}", [("Detalle", datos_t)], incluir_grafico=dict_pie)
            st.download_button("Guardar Mensual", pdf_m, f"Mensual_{p_sel}.pdf")

    if st.session_state.usuario_actual == "Admin - Ver todo":
        st.divider()
        st.subheader("🌐 Visión Global del Estudio (Desvío Trimestral)")
        hist_global = {}
        for i in range(3):
            m_c = mes - i; a_c = anio
            if m_c <= 0: m_c += 12; a_c -= 1
            df_m_g = df_p[(df_p['Fecha'].dt.month == m_c) & (df_p['Fecha'].dt.year == a_c)]
            hist_global[MESES_ES[m_c]] = df_m_g[OPERARIOS_FIJOS].sum(axis=1).groupby(df_m_g['Tarea']).sum().to_dict()
        
        tareas_g = sorted(list(set([t for m in hist_global for t in hist_global[m].keys()])))
        meses_g = list(hist_global.keys())
        header_g = ["Tarea"] + meses_g
        rows_g = []; totales_g = [0.0] * len(meses_g)
        for t in tareas_g:
            fila = [t]
            for idx, m in enumerate(meses_g):
                val = round(float(hist_global[m].get(t, 0)), 1)
                fila.append(val); totales_g[idx] += val
            rows_g.append(fila)
        rows_g.append(["TOTAL"] + [round(x, 1) for x in totales_g])
        st.table(pd.DataFrame(rows_g[1:], columns=header_g))
        
        if st.button("📥 Descargar Reporte Global Trimestral (PDF)"):
            pdf_g = generar_pdf_base("Reporte Global Trimestral", "Estudio Completo", [("Totales Equipo", [header_g] + rows_g)])
            st.download_button("Guardar Global", pdf_g, "Global_Trimestral.pdf")

# ===== 8. CARGAR HORAS E HISTORIAL =====
elif "Cargar" in menu:
    st.title("➕ Registro de Horas")
    u_c = st.session_state.usuario_actual if st.session_state.usuario_actual != "Admin - Ver todo" else st.selectbox("Persona:", OPERARIOS_FIJOS)
    
    with st.form("f_ind"):
        f_f = st.date_input("Fecha", value=datetime.now())
        f_t = st.selectbox("Tarea", list(COLORES_TAREAS.keys()))
        f_h = st.number_input("Horas", step=0.5, value=HORAS_DIA_LABORAL*1.0)
        f_n = st.text_input("Nota")
        if st.form_submit_button("Guardar Registro"):
            nueva = {'Fecha': f_f, 'Tarea': f_t, 'Nota': f_n}
            for op in OPERARIOS_FIJOS: nueva[op] = f_h if op == u_c else 0
            st.session_state.cargas = pd.concat([st.session_state.cargas, pd.DataFrame([nueva])], ignore_index=True)
            guardar_df("Cargas", st.session_state.cargas)
            st.success("¡Guardado!"); time.sleep(1); st.rerun()

    st.divider()
    st.subheader("📋 Historial y Resumen")
    df_h = st.session_state.cargas.copy(); df_h['Fecha'] = pd.to_datetime(df_h['Fecha'], errors='coerce')
    df_h = df_h[df_h[u_c] > 0]
    
    if not df_h.empty:
        # Resumen mes actual
        df_m = df_h[df_h['Fecha'].dt.month == datetime.now().month]
        if not df_m.empty:
            st.write(f"**Resumen de Tareas - {MESES_ES[datetime.now().month]}**")
            st.dataframe(df_m.groupby('Tarea')[u_c].sum().round(1).reset_index(), use_container_width=True, hide_index=True)
        
        st.write("**Últimos Registros:**")
        for i, row in df_h.sort_values('Fecha', ascending=False).head(10).iterrows():
            c1, c2 = st.columns([6, 1])
            c1.write(f"📅 {row['Fecha'].strftime('%d/%m/%Y')} | {row['Tarea']} | {row[u_c]} hs")
            if c2.button("Eliminar", key=f"del_{i}"):
                st.session_state.cargas = st.session_state.cargas.drop(i).reset_index(drop=True)
                guardar_df("Cargas", st.session_state.cargas); st.rerun()

# (Secciones Masiva, Protocolo y Reset)
elif "Carga Masiva" in menu:
    st.title("📁 Reparto de Horas (Admin)")
    with st.form("f_masiva"):
        u_m = st.selectbox("Operario", OPERARIOS_FIJOS)
        t_m = st.selectbox("Tarea", list(COLORES_TAREAS.keys()))
        f_i = st.date_input("Desde"); f_f = st.date_input("Hasta")
        h_t = st.number_input("Horas Totales", min_value=0.0)
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
    st.markdown("### Reglas\n1. Identidad propia.\n2. Carga antes de 15hs.\n3. 6 horas diarias.")
    if st.button("📥 Descargar PDF"):
        pdf = generar_pdf_base("Protocolo", "Reglas del sistema", [("Pautas", [["Regla", "Detalle"], ["Carga", "6hs diarias"]])])
        st.download_button("Guardar", pdf, "Protocolo.pdf")

elif "Reset" in menu:
    if st.text_input("Escriba BORRAR") == "BORRAR":
        if st.button("Eliminar Todo"): guardar_df("Cargas", pd.DataFrame(columns=['Fecha', 'Tarea'] + OPERARIOS_FIJOS + ['Nota'])); st.rerun()
```[cite: 1, 6]
