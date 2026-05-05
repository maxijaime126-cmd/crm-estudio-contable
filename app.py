import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import plotly.express as px
import gspread
from google.oauth2.service_account import Credentials
from io import BytesIO
import time

# ===== 1. CONFIGURACIÓN ESTRATÉGICA =====
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
TAREAS_DESCUENTO_CAPACIDAD = ["INASISTENCIA POR EXAMEN O TRAMITE"]
TAREAS_DISPONIBILIDAD = ["DISPONIBLE", "PLANIFICACIONES/ORGANIZACIÓN/PROCEDIMIENTO S/INFORMES"]
MESES_ES = {1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
            7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"}

# ===== 2. FUNCIONES PDF =====
def generar_pdf_base(titulo_doc, subtitulo, datos_tablas, incluir_grafico=None, es_protocolo=False):
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
    estilo_titulo = s['Title']; estilo_titulo.textColor = color_celeste
    estilo_cuerpo = s['Normal']; estilo_cuerpo.fontSize = 9; estilo_cuerpo.leading = 12
    estilo_subtitulo = s['Heading3']; estilo_subtitulo.textColor = color_celeste
    estilo_negrita = s['Normal']; estilo_negrita.fontSize = 9; estilo_negrita.fontName = 'Helvetica-Bold'

    story = [Paragraph("GRUPO PRESSACCO", estilo_titulo), Paragraph(titulo_doc, s['Heading2']), Paragraph(subtitulo, s['Normal']), Spacer(1, 15)]

    if es_protocolo:
        secciones = [
            ("1. INTRODUCCIÓN Y FINALIDAD", "El objetivo es transformar nuestra carga de trabajo en datos accionables. El tiempo permite medir rentabilidad y planificar el crecimiento.[cite: 1]"),
            ("2. OBJETIVOS ESTRATÉGICOS", "• Visibilidad de tareas. • Equilibrio de cargas. • Transparencia histórica.[cite: 1]"),
            ("3. PASO A PASO: CARGA Y CONTROL", "• Registro Diario: 6 horas antes de las 15:00 hs. • Inasistencias: Se cargan para restar capacidad neta.[cite: 1]"),
            ("4. SEMÁFORO DE CAPACIDAD", "• 🟢 Verde (>20%). • 🟡 Amarillo (10-20%). • 🔴 Rojo (<10%).[cite: 1]")
        ]
        for t, c in secciones:
            story.append(Paragraph(t, estilo_subtitulo)); story.append(Paragraph(c, estilo_cuerpo)); story.append(Spacer(1, 10))

    if incluir_grafico:
        d = Drawing(450, 200); pc = Pie(); pc.x = 50; pc.y = 25; pc.width = 130; pc.height = 130
        lista_colores = [colors.magenta, colors.deepskyblue, colors.lightpink, colors.yellow, colors.whitesmoke, colors.lightblue, colors.lavender, colors.bisque, colors.turquoise, colors.lime, colors.hotpink]
        grafico_limpio = {k: v for k, v in incluir_grafico.items() if k not in TAREAS_DESCUENTO_CAPACIDAD}
        total_h = sum(grafico_limpio.values())
        pc.data = [round(float(v), 1) for v in grafico_limpio.values()]
        pc.labels = [f"{round((v/total_h)*100, 1)}%" if total_h > 0 else "0%" for v in grafico_limpio.values()]
        for i in range(len(pc.data)): pc.slices[i].fillColor = lista_colores[i % len(lista_colores)]
        leg = Legend(); leg.x = 220; leg.y = 150; leg.alignment = 'right'; leg.columnMaximum = 12; leg.fontSize = 7
        leg.colorNamePairs = [(lista_colores[i % len(lista_colores)], k) for i, k in enumerate(grafico_limpio.keys())]
        d.add(pc); d.add(leg); story.append(d); story.append(Spacer(1, 10))

    for titulo_tabla, data in datos_tablas:
        if titulo_tabla: story.append(Paragraph(titulo_tabla, s['Heading3']))
        data_p = [[Paragraph(str(c), estilo_negrita if "TOTAL" in str(c).upper() else estilo_cuerpo) for c in fila] for fila in data]
        col_w = [2.5*inch] + [1.0*inch]*(len(data[0])-1)
        t = Table(data_p, colWidths=col_w)
        estilo_t = [('BACKGROUND', (0, 0), (-1, 0), color_celeste), ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke), ('GRID', (0, 0), (-1, -1), 0.5, colors.grey), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')]
        if any("TOTAL" in str(c).upper() for c in data[-1]): estilo_t.append(('BACKGROUND', (0, -1), (-1, -1), colors.lightgrey))
        t.setStyle(TableStyle(estilo_t)); story.append(t); story.append(Spacer(1, 15))

    doc.build(story); buf.seek(0); return buf

# ===== 3. CONEXIÓN =====
@st.cache_resource
def conectar():
    creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=["https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/drive"])
    return gspread.authorize(creds).open("CRM_Estudio_Datos")

def cargar_hoja(nombre):
    try:
        ws = conectar().worksheet(nombre); df = pd.DataFrame(ws.get_all_records())
        for c in df.columns:
            if 'fecha' in c.lower(): df[c] = pd.to_datetime(df[c], dayfirst=True, errors='coerce')
        return df, ws
    except: return pd.DataFrame(), None

def guardar_df(nombre, df):
    try:
        ws = conectar().worksheet(nombre); ws.clear(); df_c = df.copy()
        for c in df_c.columns:
            if 'fecha' in c.lower(): df_c[c] = pd.to_datetime(df_c[c], errors='coerce').dt.strftime('%d/%m/%Y')
        ws.update([df_c.columns.values.tolist()] + df_c.fillna('').astype(str).values.tolist())
        st.cache_data.clear(); return True
    except: return False

@st.cache_data
def cargar_feriados():
    try:
        df_f = pd.read_csv("feriados_2026.csv"); df_f.columns = df_f.columns.str.strip().str.lower()
        return pd.to_datetime(df_f['fecha'], errors='coerce').dt.date.dropna().tolist()
    except: return []

FERIADOS = cargar_feriados()

if 'cargas' not in st.session_state:
    df, _ = cargar_hoja("Cargas"); st.session_state.cargas = df if not df.empty else pd.DataFrame(columns=['Fecha', 'Tarea'] + OPERARIOS_FIJOS + ['Nota'])

if 'usuario_actual' not in st.session_state: st.session_state.usuario_actual = None

# ===== 4. LOGIN Y NAVEGACIÓN =====
if st.session_state.usuario_actual is None:
    st.title("🏛️ CRM Grupo Pressacco")
    u = st.selectbox("Usuario:", ["Seleccionar..."] + OPERARIOS_FIJOS + ["Admin - Ver todo"])
    if st.button("Ingresar") and u != "Seleccionar...": st.session_state.usuario_actual = u; st.rerun()
    st.stop()

def mostrar_alerta_mensual(usuario):
    if usuario == "Admin - Ver todo" or usuario is None: return
    hoy = datetime.now()
    ini_m = datetime(hoy.year, hoy.month, 1).date()
    fin_m = (datetime(hoy.year, hoy.month + 1, 1) if hoy.month < 12 else datetime(hoy.year + 1, 1, 1)).date() - timedelta(days=1)
    dias_h = len(pd.bdate_range(start=ini_m, end=fin_m, freq='C', holidays=FERIADOS))
    total_obj = dias_h * HORAS_DIA_LABORAL
    df_u = st.session_state.cargas.copy(); df_u['Fecha'] = pd.to_datetime(df_u['Fecha'], errors='coerce')
    cargadas = df_u[(df_u['Fecha'].dt.month == hoy.month) & (df_u['Fecha'].dt.year == hoy.year)][usuario].sum()
    restante = total_obj - cargadas
    if restante > 0:
        st.warning(f"🎯 **Objetivo {MESES_ES[hoy.month]}:** Faltan **{round(restante, 1)} hs** para las {total_obj} hs del mes.[cite: 1]")
    else:
        st.success(f"✅ ¡Objetivo de {total_obj} hs cumplido!")

mostrar_alerta_mensual(st.session_state.usuario_actual)

menu = st.sidebar.radio("Navegación", ["📊 Panel de Control", "➕ Cargar Horas", "📁 Carga Masiva", "📜 Protocolo", "⚙️ Reset"] if st.session_state.usuario_actual == "Admin - Ver todo" else ["📊 Panel de Control", "➕ Cargar Mis Horas", "📜 Protocolo"])
if st.sidebar.button("Cerrar Sesión"): st.session_state.clear(); st.rerun()

# ===== 5. PANEL DE CONTROL =====
if "Panel de Control" in menu:
    st.title("📊 Análisis de Capacidad y Eficiencia")
    c1, c2, c3 = st.columns([1,1,2])
    with c1: anio = st.selectbox("Año", [2025, 2026], index=1)
    with c2: mes = st.selectbox("Mes", list(range(1,13)), format_func=lambda x: MESES_ES[x], index=datetime.now().month-1)
    with c3: p_sel = st.selectbox("Integrante Individual:", OPERARIOS_FIJOS) if st.session_state.usuario_actual == "Admin - Ver todo" else st.session_state.usuario_actual

    df_p = st.session_state.cargas.copy(); df_p['Fecha'] = pd.to_datetime(df_p['Fecha'], errors='coerce')
    
    st.subheader(f"📈 Comparativa Trimestral - {p_sel}")
    comp_list = []; hist_pdf = {}
    for i in range(3):
        m_c = mes - i; a_c = anio
        if m_c <= 0: m_c += 12; a_c -= 1
        ini_c = datetime(a_c, m_c, 1).date()
        fin_c = (datetime(a_c, m_c + 1, 1) if m_c < 12 else datetime(a_c + 1, 1, 1)).date() - timedelta(days=1)
        cap_t = len(pd.bdate_range(start=ini_c, end=fin_c, freq='C', holidays=FERIADOS)) * HORAS_DIA_LABORAL
        df_m = df_p[(df_p['Fecha'].dt.month == m_c) & (df_p['Fecha'].dt.year == a_c)]
        h_ina = df_m[df_m['Tarea'].isin(TAREAS_DESCUENTO_CAPACIDAD)][p_sel].sum()
        cap_n = cap_t - h_ina
        total_b = round(df_m[p_sel].sum(), 1)
        h_disp = df_m[df_m['Tarea'].isin(TAREAS_DISPONIBILIDAD)][p_sel].sum()
        disp_v = (h_disp / cap_n * 100) if cap_n > 0 else 0
        semaforo = "🟢 (Libre)" if disp_v > 20 else "🟡 (Atención)" if disp_v >= 10 else "🔴 (Saturado)"
        comp_list.append({"Mes": MESES_ES[m_c], "Carga": f"{total_b} hs", "Cap. Neta": f"{cap_n} hs", "Disponibilidad": f"{round(disp_v, 1)}%", "Estado": semaforo})
        hist_pdf[MESES_ES[m_c]] = df_m.groupby('Tarea')[p_sel].sum().to_dict()

    st.table(pd.DataFrame(comp_list))
    
    if st.button("📥 PDF Trimestral"):
        tareas_u = sorted(list(set([t for m in hist_pdf for t in hist_pdf[m].keys()]))); meses_n = list(hist_pdf.keys())
        rows = []; tot_m = [0.0] * len(meses_n)
        for t in tareas_u:
            f = [t]
            for idx, m in enumerate(meses_n):
                v = round(float(hist_pdf[m].get(t, 0)), 1); f.append(v); tot_m[idx] += v
            rows.append(f)
        rows.append(["TOTAL BRUTO"] + [round(x, 1) for x in tot_m])
        st.download_button("Guardar Trimestral", generar_pdf_base(f"Trimestral: {p_sel}", "Resumen 3 meses", [("Detalle", [["Tarea"] + meses_n] + rows)]), f"Trimestral_{p_sel}.pdf")

    st.divider()
    # GRÁFICO INDIVIDUAL (CORREGIDO PARA TODOS LOS USUARIOS)
    df_act = df_p[(df_p['Fecha'].dt.month == mes) & (df_p['Fecha'].dt.year == anio)]
    res_ind = df_act.groupby('Tarea')[p_sel].sum().round(1).reset_index()
    res_graf = res_ind[(res_ind[p_sel] > 0) & (~res_ind['Tarea'].isin(TAREAS_DESCUENTO_CAPACIDAD))]
    
    if not res_graf.empty:
        col_g, col_m = st.columns([2,1])
        with col_g:
            fig = px.pie(res_graf, values=p_sel, names='Tarea', color='Tarea', color_discrete_map=COLORES_TAREAS, hole=0.5, title=f"Eficiencia Real - {p_sel}")
            fig.update_traces(textposition='inside', textinfo='percent+label')
            st.plotly_chart(fig, use_container_width=True)
        with col_m:
            st.metric("Disponibilidad", comp_list[0]["Disponibilidad"])
            st.metric("Estado", comp_list[0]["Estado"])
            if st.button("📥 PDF Mensual"):
                dat = [["Tarea", "Horas"]] + [[r['Tarea'], r[p_sel]] for _, r in res_ind.iterrows()] + [["TOTAL", comp_list[0]["Carga"]]]
                st.download_button("Guardar Mensual", generar_pdf_base(f"Reporte {p_sel}", f"{MESES_ES[mes]} {anio}", [("Detalle", dat)], incluir_grafico=res_ind.set_index('Tarea')[p_sel].to_dict()), f"Mensual_{p_sel}.pdf")

    # VISIÓN GLOBAL (ADMIN)
    if st.session_state.usuario_actual == "Admin - Ver todo":
        st.divider(); st.subheader("🌐 Visión Global del Estudio (Equipo)")
        df_eq = df_act[~df_act['Tarea'].isin(TAREAS_DESCUENTO_CAPACIDAD)].melt(id_vars=['Fecha', 'Tarea'], value_vars=OPERARIOS_FIJOS, var_name='Op', value_name='Hs')
        res_eq = df_eq.groupby('Tarea')['Hs'].sum().reset_index()
        fig_g = px.pie(res_eq, values='Hs', names='Tarea', color='Tarea', color_discrete_map=COLORES_TAREAS, hole=0.5, title="Total Horas Equipo")
        fig_g.update_traces(textposition='inside', textinfo='percent+label')
        st.plotly_chart(fig_g, use_container_width=True)
        
        if st.button("📥 PDF Global"):
            hist_g = {}
            for i in range(3):
                m_c = mes - i; a_c = anio
                if m_c <= 0: m_c += 12; a_c -= 1
                df_m_g = df_p[(df_p['Fecha'].dt.month == m_c) & (df_p['Fecha'].dt.year == a_c)]
                hist_g[MESES_ES[m_c]] = df_m_g[~df_m_g['Tarea'].isin(TAREAS_DESCUENTO_CAPACIDAD)][OPERARIOS_FIJOS].sum(axis=1).groupby(df_m_g['Tarea']).sum().to_dict()
            tg = sorted(list(set([t for m in hist_g for t in hist_g[m].keys()]))); mg = list(hist_g.keys())
            rg = []; totg = [0.0] * len(mg)
            for t in tg:
                f = [t]; 
                for idx, m in enumerate(mg):
                    v = round(float(hist_g[m].get(t, 0)), 1); f.append(v); totg[idx] += v
                rg.append(f)
            rg.append(["TOTAL NETO"] + [round(x, 1) for x in totg])
            st.download_button("Guardar Global", generar_pdf_base("REPORTE GLOBAL", "Estudio Completo", [("Consolidado", [["Tarea"] + mg] + rg)], incluir_grafico=res_eq.set_index('Tarea')['Hs'].to_dict()), "Global_Neto.pdf")

# (Secciones Carga Masiva, Carga Individual, Protocolo y Reset intactas)
elif "Carga Masiva" in menu:
    st.title("📁 Distribución Masiva")
    with st.form("fm"):
        u_m = st.selectbox("Persona", OPERARIOS_FIJOS); t_m = st.selectbox("Tarea", list(COLORES_TAREAS.keys()))
        f_i = st.date_input("Desde"); f_f = st.date_input("Hasta"); h_t = st.number_input("Horas Totales", min_value=0.0)
        if st.form_submit_button("Ejecutar"):
            rgo = pd.bdate_range(start=f_i, end=f_f, freq='C', holidays=FERIADOS)
            if len(rgo) > 0:
                h_d = round(h_t / len(rgo), 2); filas = []
                for d in rgo:
                    fila = {'Fecha': d, 'Tarea': t_m, 'Nota': 'Masiva'}; 
                    for o in OPERARIOS_FIJOS: fila[o] = h_d if o == u_m else 0
                    filas.append(fila)
                st.session_state.cargas = pd.concat([st.session_state.cargas, pd.DataFrame(filas)], ignore_index=True)
                if guardar_df("Cargas", st.session_state.cargas):
                    st.success(f"✅ Guardado con éxito."); time.sleep(1.5); st.rerun()

elif "Cargar" in menu:
    st.title("➕ Registro de Horas")
    u_c = st.session_state.usuario_actual if st.session_state.usuario_actual != "Admin - Ver todo" else st.selectbox("Persona:", OPERARIOS_FIJOS)
    with st.form("fi"):
        f_f = st.date_input("Fecha", value=datetime.now()); f_t = st.selectbox("Tarea", list(COLORES_TAREAS.keys())); f_h = st.number_input("Horas", step=0.5, value=6.0); f_n = st.text_input("Nota")
        if st.form_submit_button("Guardar"):
            nueva = {'Fecha': f_f, 'Tarea': f_t, 'Nota': f_n}
            for op in OPERARIOS_FIJOS: nueva[op] = f_h if op == u_c else 0
            st.session_state.cargas = pd.concat([st.session_state.cargas, pd.DataFrame([nueva])], ignore_index=True)
            guardar_df("Cargas", st.session_state.cargas); st.success("¡Guardado!"); time.sleep(1); st.rerun()
    st.divider(); mes_f = st.selectbox("Historial Mes:", list(range(1,13)), format_func=lambda x: MESES_ES[x], index=datetime.now().month-1)
    df_h = st.session_state.cargas.copy(); df_h['Fecha'] = pd.to_datetime(df_h['Fecha'], errors='coerce')
    df_f = df_h[(df_h[u_c] > 0) & (df_h['Fecha'].dt.month == mes_f)]
    if not df_f.empty:
        st.info(f"**Total {MESES_ES[mes_f]}:** {round(df_f[u_c].sum(), 1)} hs")
        for i, r in df_f.sort_values('Fecha', ascending=False).iterrows():
            c1, c2 = st.columns([6, 1])
            c1.write(f"📅 {r['Fecha'].strftime('%d/%m/%Y')} | {r['Tarea']} | {r[u_c]} hs | {r['Nota']}")
            if c2.button("Eliminar", key=f"del_{i}"): st.session_state.cargas.drop(i, inplace=True); guardar_df("Cargas", st.session_state.cargas); st.rerun()

elif "Protocolo" in menu:
    st.title("📜 Protocolo de Uso")
    if st.button("📥 Descargar Guía"):
        st.download_button("Guardar", generar_pdf_base("PROTOCOLO CRM", "Manual de Procedimientos", [], es_protocolo=True), "Protocolo_Pressacco.pdf")

elif "Reset" in menu:
    if st.text_input("Escriba BORRAR") == "BORRAR":
        if st.button("Limpiar Todo"): guardar_df("Cargas", pd.DataFrame(columns=['Fecha', 'Tarea'] + OPERARIOS_FIJOS + ['Nota'])); st.rerun()
