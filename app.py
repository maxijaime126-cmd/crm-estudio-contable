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
MESES_ES = {1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
            7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"}

# ===== 2. FUNCIONES PDF MEJORADAS =====

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
    estilo_titulo = s['Title']
    estilo_titulo.textColor = color_celeste
    estilo_cuerpo = s['Normal']
    estilo_cuerpo.fontSize = 9
    estilo_negrita = s['Normal']
    estilo_negrita.fontSize = 9
    estilo_negrita.fontName = 'Helvetica-Bold'

    story = [
        Paragraph("GRUPO PRESSACCO", estilo_titulo),
        Paragraph(titulo_doc, s['Heading2']),
        Paragraph(subtitulo, s['Normal']),
        Spacer(1, 15)
    ]

    # Contenido extendido para el Protocolo dentro del PDF
    if es_protocolo:
        contenido = [
            ("<b>¿Para qué sirve este CRM?</b>", "Para medir nuestra capacidad real y eficiencia. Si no registramos las horas, no sabemos cuánto nos lleva cada cliente o tarea. Con este sistema, tenemos la prueba de nuestra carga de trabajo, detectamos cuellos de botella y podemos planificar mejor el mes."),
            ("<b>INGRESO Y CARGA - El registro diario</b>", "Acá es donde alimentamos el sistema. Sin carga diaria, los gráficos dan error.<br/>• <b>Usuario:</b> Obligatorio entrar con tu nombre. Nunca cargues en el de un compañero.<br/>• <b>Tarea:</b> Elegí la que realmente hiciste.<br/>• <b>Horas:</b> El total del día debe ser siempre de 6 horas.<br/>• <b>Nota:</b> Escribí brevemente qué hiciste."),
            ("<b>PANEL DE CONTROL - El espejo del equipo</b>", "Acá analizamos resultados. Podés filtrar por mes y descargar tus reportes mensuales o trimestrales para las reuniones."),
            ("<b>REGLAS DE ORO</b>", "• Carga antes de las 15:00 hs.<br/>• Nunca toques el Google Sheets a mano.<br/>• Sinceridad: Si no tuviste tareas, cargá DISPONIBLE."),
            ("<b>FLUJO DE TRABAJO DIARIO</b>", "1. Al ingresar mirá el cartel de aviso.<br/>2. Durante el día registrá tareas importantes.<br/>3. Al cierre verificá que sumen 6 horas.")
        ]
        for t, c in contenido:
            story.append(Paragraph(t, s['Heading3']))
            story.append(Paragraph(c, estilo_cuerpo))
            story.append(Spacer(1, 10))

    if incluir_grafico:
        d = Drawing(450, 200)
        pc = Pie()
        pc.x = 50; pc.y = 25; pc.width = 130; pc.height = 130
        lista_colores = [colors.magenta, colors.deepskyblue, colors.lightpink, colors.yellow, colors.whitesmoke, colors.lightblue, colors.lavender, colors.bisque, colors.turquoise, colors.lime, colors.hotpink]
        total_h = sum(incluir_grafico.values())
        pc.data = [round(float(v), 1) for v in incluir_grafico.values()]
        pc.labels = [f"{round((v/total_h)*100, 1)}%" if total_h > 0 else "0%" for v in incluir_grafico.values()]
        for i in range(len(pc.data)): pc.slices[i].fillColor = lista_colores[i % len(lista_colores)]
        leg = Legend()
        leg.x = 220; leg.y = 150; leg.alignment = 'right'; leg.columnMaximum = 12; leg.fontSize = 7
        leg.colorNamePairs = [(lista_colores[i % len(lista_colores)], k) for i, k in enumerate(incluir_grafico.keys())]
        d.add(pc); d.add(leg); story.append(d); story.append(Spacer(1, 10))

    for titulo_tabla, data in datos_tablas:
        if titulo_tabla: story.append(Paragraph(titulo_tabla, s['Heading3']))
        data_p = [[Paragraph(str(c), estilo_negrita if str(fila[0]).upper() == "TOTAL" else estilo_cuerpo) for c in fila] for fila in data]
        col_w = [2.5*inch] + [1.0*inch]*(len(data[0])-1)
        t = Table(data_p, colWidths=col_w)
        estilo_t = [('BACKGROUND', (0, 0), (-1, 0), color_celeste), ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke), ('GRID', (0, 0), (-1, -1), 0.5, colors.grey), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')]
        if str(data[-1][0]).upper() == "TOTAL": estilo_t.append(('BACKGROUND', (0, -1), (-1, -1), colors.lightgrey))
        t.setStyle(TableStyle(estilo_t))
        story.append(t); story.append(Spacer(1, 15))

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

if 'usuario_actual' not in st.session_state: st.session_state.usuario_actual = None

# ===== 4. LÓGICA DE ALERTA =====
def mostrar_alerta_faltante(usuario):
    if usuario == "Admin - Ver todo" or usuario is None: return
    hoy = datetime.now()
    inicio_mes = datetime(hoy.year, hoy.month, 1).date()
    dias_habiles = len(pd.bdate_range(start=inicio_mes, end=hoy.date(), freq='C', holidays=FERIADOS))
    horas_obj = dias_habiles * HORAS_DIA_LABORAL
    df_u = st.session_state.cargas.copy(); df_u['Fecha'] = pd.to_datetime(df_u['Fecha'], errors='coerce')
    horas_c = df_u[(df_u['Fecha'].dt.month == hoy.month) & (df_u['Fecha'].dt.year == hoy.year)][usuario].sum()
    if horas_c < horas_obj:
        st.warning(f"⚠️ **Aviso:** Hola {usuario}, te faltan cargar **{round(horas_obj - horas_c, 1)} horas** para completar el mes.")

# ===== 5. LOGIN =====
if st.session_state.usuario_actual is None:
    st.title("🏛️ CRM Grupo Pressacco")
    u = st.selectbox("Usuario:", ["Seleccionar..."] + OPERARIOS_FIJOS + ["Admin - Ver todo"])
    if st.button("Ingresar") and u != "Seleccionar...":
        st.session_state.usuario_actual = u; st.rerun()
    st.stop()

# ===== 6. NAVEGACIÓN =====
opciones = ["📊 Panel de Control", "➕ Cargar Horas", "📁 Carga Masiva", "📜 Protocolo", "⚙️ Reset"] if st.session_state.usuario_actual == "Admin - Ver todo" else ["📊 Panel de Control", "➕ Cargar Mis Horas", "📜 Protocolo"]
menu = st.sidebar.radio("Navegación", opciones)
if st.sidebar.button("Cerrar Sesión"): st.session_state.clear(); st.rerun()

mostrar_alerta_faltante(st.session_state.usuario_actual)

# ===== 7. SECCIONES =====
if "Panel de Control" in menu:
    st.title("📊 Análisis de Ocupación")
    c1, c2, c3 = st.columns([1,1,2])
    with c1: anio = st.selectbox("Año", [2025, 2026, 2027], index=1)
    with c2: mes = st.selectbox("Mes", list(range(1,13)), format_func=lambda x: MESES_ES[x], index=datetime.now().month-1)
    with c3: p_sel = st.selectbox("Integrante:", OPERARIOS_FIJOS) if st.session_state.usuario_actual == "Admin - Ver todo" else st.session_state.usuario_actual

    df_p = st.session_state.cargas.copy(); df_p['Fecha'] = pd.to_datetime(df_p['Fecha'], errors='coerce')
    
    st.subheader(f"📈 Comparativa Trimestral - {p_sel}")
    hist_pdf = {}
    for i in range(3):
        m_c = mes - i; a_c = anio
        if m_c <= 0: m_c += 12; a_c -= 1
        df_m = df_p[(df_p['Fecha'].dt.month == m_c) & (df_p['Fecha'].dt.year == a_c)]
        hist_pdf[MESES_ES[m_c]] = df_m.groupby('Tarea')[p_sel].sum().to_dict()
    
    # PDF Trimestral con Totales
    if st.button("📥 Descargar Autoevaluación Trimestral (PDF)"):
        tareas_u = sorted(list(set([t for m in hist_pdf for t in hist_pdf[m].keys()])))
        meses_n = list(hist_pdf.keys())
        rows = []; totales_m = [0.0] * len(meses_n)
        for t in tareas_u:
            fila = [t]
            for idx, m in enumerate(meses_n):
                val = round(float(hist_pdf[m].get(t, 0)), 1); fila.append(val); totales_m[idx] += val
            rows.append(fila)
        rows.append(["TOTAL"] + [round(x, 1) for x in totales_m])
        pdf_t = generar_pdf_base(f"Autoevaluación Trimestral: {p_sel}", "Comparativa mensual", [("Desvío por Tarea", [["Tarea"] + meses_n] + rows)])
        st.download_button("Guardar Trimestral", pdf_t, f"Trimestral_{p_sel}.pdf")

    st.divider()
    df_act = df_p[(df_p['Fecha'].dt.month == mes) & (df_p['Fecha'].dt.year == anio)]
    res_ind = df_act.groupby('Tarea')[p_sel].sum().round(1).reset_index()
    res_ind = res_ind[res_ind[p_sel]>0]
    
    if not res_ind.empty:
        col_p, col_m = st.columns([2,1])
        with col_p: st.plotly_chart(px.pie(res_ind, values=p_sel, names='Tarea', color='Tarea', color_discrete_map=COLORES_TAREAS), use_container_width=True)
        with col_m:
            if st.button("📥 Descargar Reporte Mensual (PDF)"):
                tot_h = res_ind[p_sel].sum()
                datos_t = [["Tarea", "Horas", "%"]] + [[r['Tarea'], round(r[p_sel], 1), f"{round((r[p_sel]/tot_h)*100, 1)}%"] for _, r in res_ind.iterrows()] + [["TOTAL", round(tot_h, 1), "100%"]]
                pdf_m = generar_pdf_base(f"Reporte Mensual: {p_sel}", f"{MESES_ES[mes]} {anio}", [("Detalle", datos_t)], incluir_grafico=res_ind.set_index('Tarea')[p_sel].to_dict())
                st.download_button("Guardar Mensual", pdf_m, f"Mensual_{p_sel}.pdf")

elif "Protocolo" in menu:
    st.title("📜 Protocolo de Uso - Grupo Pressacco")
    st.markdown("### ¿Para qué sirve este CRM?[cite: 6]")
    st.write("Para medir capacidad real y eficiencia. Detectar cuellos de botella y planificar mejor el mes.[cite: 6]")
    
    if st.button("📥 Descargar Guía Maestra (PDF)"):
        # PDF del Protocolo con contenido completo adentro[cite: 6]
        pdf_p = generar_pdf_base("PROTOCOLO DE USO - CRM", "Guía Completa de Funcionamiento", [], es_protocolo=True)
        st.download_button("Guardar Protocolo", pdf_p, "Protocolo_Pressacco.pdf")

elif "Cargar" in menu:
    st.title("➕ Registro de Horas")
    u_c = st.session_state.usuario_actual if st.session_state.usuario_actual != "Admin - Ver todo" else st.selectbox("Persona:", OPERARIOS_FIJOS)
    with st.form("f_ind"):
        f_f = st.date_input("Fecha", value=datetime.now()); f_t = st.selectbox("Tarea", list(COLORES_TAREAS.keys())); f_h = st.number_input("Horas", step=0.5, value=6.0); f_n = st.text_input("Nota")
        if st.form_submit_button("Guardar Registro"):
            nueva = {'Fecha': f_f, 'Tarea': f_t, 'Nota': f_n}
            for op in OPERARIOS_FIJOS: nueva[op] = f_h if op == u_c else 0
            st.session_state.cargas = pd.concat([st.session_state.cargas, pd.DataFrame([nueva])], ignore_index=True)
            guardar_df("Cargas", st.session_state.cargas); st.success("¡Guardado!"); time.sleep(1); st.rerun()

    st.divider(); st.subheader("📋 Historial y Consulta")
    mes_filt = st.selectbox("Consultar Mes:", list(range(1,13)), format_func=lambda x: MESES_ES[x], index=datetime.now().month-1)
    df_h = st.session_state.cargas.copy(); df_h['Fecha'] = pd.to_datetime(df_h['Fecha'], errors='coerce')
    df_f = df_h[(df_h[u_c] > 0) & (df_h['Fecha'].dt.month == mes_filt)]
    if not df_f.empty:
        st.dataframe(df_f.groupby('Tarea')[u_c].sum().round(1).reset_index(), use_container_width=True, hide_index=True)
        for i, row in df_f.sort_values('Fecha', ascending=False).iterrows():
            c1, c2 = st.columns([6, 1])
            c1.write(f"📅 {row['Fecha'].strftime('%d/%m/%Y')} | {row['Tarea']} | {row[u_c]} hs | {row['Nota']}")
            if c2.button("Eliminar", key=f"del_{i}"):
                st.session_state.cargas = st.session_state.cargas.drop(i).reset_index(drop=True)
                guardar_df("Cargas", st.session_state.cargas); st.rerun()
