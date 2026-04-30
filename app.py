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
TAREAS_DISPONIBILIDAD_REAL = ["DISPONIBLE", "PLANIFICACIONES/ORGANIZACIÓN/PROCEDIMIENTO S/INFORMES"]
TAREAS_DESCUENTO_CAPACIDAD = ["INASISTENCIA POR EXAMEN O TRAMITE"]
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
            ("1. INTRODUCCIÓN Y FINALIDAD", "El objetivo principal de este CRM es transformar nuestra carga de trabajo en datos accionables. En un estudio contable, el tiempo es nuestro recurso más valioso; registrarlo nos permite medir la rentabilidad de cada proceso, detectar cuándo un integrante está saturado y planificar el crecimiento del equipo con bases sólidas. Elegimos sumar precisión para restar incertidumbre.[cite: 2]"),
            ("2. OBJETIVOS ESTRATÉGICOS", "• <b>Visibilidad:</b> Saber en qué tareas invertimos más tiempo.<br/>• <b>Equilibrio:</b> Evitar cuellos de botella y redistribuir tareas.<br/>• <b>Transparencia:</b> Tener un registro histórico claro frente a auditorías.[cite: 2]"),
            ("3. PASO A PASO: CARGA Y CONTROL", "<b>• Carga Diaria:</b> Registrar 6 horas diarias antes de las 15:00 hs.<br/><b>• Sinceridad:</b> Usar tareas reales o DISPONIBLE según corresponda.<br/><b>• Inasistencias:</b> Se cargan para restar capacidad neta automáticamente.<br/><b>• Autocontrol:</b> Verificar en el Historial que el total sume 6 hs.<br/><b>• Reunión de Equipo:</b> Analizar el PDF Trimestral grupalmente.[cite: 2]"),
            ("4. SEMÁFORO DE CAPACIDAD", "• 🟢 <b>Verde (>20%):</b> Espacio para nuevos proyectos.<br/>• 🟡 <b>Amarillo (10-20%):</b> Carga próxima al límite.<br/>• 🔴 <b>Rojo (<10%):</b> Saturación operativa.[cite: 2]")
        ]
        for t, c in secciones:
            story.append(Paragraph(t, estilo_subtitulo))
            story.append(Paragraph(c, estilo_cuerpo))
            story.append(Spacer(1, 10))

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
        df_f['fecha'] = pd.to_datetime(df_f['fecha'], errors='coerce')
        return df_f['fecha'].dt.date.dropna().tolist()
    except: return []

FERIADOS = cargar_feriados()

if 'cargas' not in st.session_state:
    df, _ = cargar_hoja("Cargas")
    st.session_state.cargas = df if not df.empty else pd.DataFrame(columns=['Fecha', 'Tarea'] + OPERARIOS_FIJOS + ['Nota'])

if 'usuario_actual' not in st.session_state: st.session_state.usuario_actual = None

# ===== 4. LOGIN Y NAVEGACIÓN (DEFINICIÓN DE 'MENU') =====
if st.session_state.usuario_actual is None:
    st.title("🏛️ CRM Grupo Pressacco")
    u = st.selectbox("Usuario:", ["Seleccionar..."] + OPERARIOS_FIJOS + ["Admin - Ver todo"])
    if st.button("Ingresar") and u != "Seleccionar...":
        st.session_state.usuario_actual = u; st.rerun()
    st.stop()

# Aquí se define 'menu' ANTES de usarlo en los condicionales
opciones = ["📊 Panel de Control", "➕ Cargar Horas", "📁 Carga Masiva", "📜 Protocolo", "⚙️ Reset"] if st.session_state.usuario_actual == "Admin - Ver todo" else ["📊 Panel de Control", "➕ Cargar Mis Horas", "📜 Protocolo"]
menu = st.sidebar.radio("Navegación", opciones)

if st.sidebar.button("Cerrar Sesión"):
    st.session_state.clear(); st.rerun()

# ===== 5. LÓGICA DE ALERTA =====
def mostrar_alerta_faltante(usuario):
    if usuario == "Admin - Ver todo" or usuario is None: return
    hoy = datetime.now(); inicio_mes = datetime(hoy.year, hoy.month, 1).date()
    dias_habiles = len(pd.bdate_range(start=inicio_mes, end=hoy.date(), freq='C', holidays=FERIADOS))
    horas_obj = dias_habiles * HORAS_DIA_LABORAL
    df_u = st.session_state.cargas.copy(); df_u['Fecha'] = pd.to_datetime(df_u['Fecha'], errors='coerce')
    horas_c = df_u[(df_u['Fecha'].dt.month == hoy.month) & (df_u['Fecha'].dt.year == hoy.year)][usuario].sum()
    if horas_c < horas_obj:
        st.warning(f"⚠️ **Aviso:** Hola {usuario}, te faltan cargar **{round(horas_obj - horas_c, 1)} horas** para completar el mes.[cite: 1]")

mostrar_alerta_faltante(st.session_state.usuario_actual)

# ===== 6. SECCIONES DEL SISTEMA =====

if "Panel de Control" in menu:
    st.title("📊 Análisis y Autoevaluación")
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
        ini = datetime(a_c, m_c, 1).date(); fin = (datetime(a_c, m_c+1, 1) if m_c < 12 else datetime(a_c+1, 1, 1)).date() - timedelta(days=1)
        cap_t = len(pd.bdate_range(start=ini, end=fin, freq='C', holidays=FERIADOS)) * HORAS_DIA_LABORAL
        df_m = df_p[(df_p['Fecha'].dt.month == m_c) & (df_p['Fecha'].dt.year == a_c)]
        h_inasist = df_m[df_m['Tarea'].isin(TAREAS_DESCUENTO_CAPACIDAD)][p_sel].sum()
        cap_n = cap_t - h_inasist
        total_real = round(df_m[p_sel].sum(), 1)
        total_neto = round(total_real - h_inasist, 1)
        h_disp_p = df_m[df_m['Tarea'].isin(TAREAS_DISPONIBILIDAD_REAL)][p_sel].sum()
        dispon_v = (h_disp_p / cap_n * 100) if cap_n > 0 else 0
        semaforo = "🟢 (Libre)" if dispon_v > 20 else "🟡 (Atención)" if dispon_v >= 10 else "🔴 (Preocupación)"
        comp_list.append({"Mes": MESES_ES[m_c], "Carga Total": f"{total_real} hs", "Total Neto": f"{total_neto} hs", "Disponibilidad": f"{dispon_v:.1f}%", "Estado": semaforo})
        hist_pdf[MESES_ES[m_c]] = df_m.groupby('Tarea')[p_sel].sum().to_dict()

    st.table(pd.DataFrame(comp_list))
    
    if st.button("📥 Descargar Autoevaluación Trimestral (PDF)"):
        tareas_u = sorted(list(set([t for m in hist_pdf for t in hist_pdf[m].keys()]))); meses_n = list(hist_pdf.keys())
        rows = []; totales_m = [0.0] * len(meses_n)
        for t in tareas_u:
            fila = [t]
            for idx, m in enumerate(meses_n):
                val = round(float(hist_pdf[m].get(t, 0)), 1); fila.append(val); totales_m[idx] += val
            rows.append(fila)
        rows.append(["TOTAL BRUTO"] + [round(x, 1) for x in totales_m])
        pdf_t = generar_pdf_base(f"Trimestral Detallado: {p_sel}", "Comparativa de horas registradas[cite: 1]", [("Desvío por Tarea", [["Tarea"] + meses_n] + rows)])
        st.download_button("Guardar Trimestral", pdf_t, f"Trimestral_{p_sel}.pdf")

    st.divider()
    df_act = df_p[(df_p['Fecha'].dt.month == mes) & (df_p['Fecha'].dt.year == anio)]
    res_ind = df_act.groupby('Tarea')[p_sel].sum().round(1).reset_index()
    res_ind = res_ind[res_ind[p_sel]>0]
    
    if not res_ind.empty:
        col_p, col_m = st.columns([2,1])
        res_neta_grafico = res_ind[~res_ind['Tarea'].isin(TAREAS_DESCUENTO_CAPACIDAD)]
        with col_p: st.plotly_chart(px.pie(res_neta_grafico, values=p_sel, names='Tarea', color='Tarea', color_discrete_map=COLORES_TAREAS, title=f"Eficiencia Real - {p_sel}"), use_container_width=True)
        with col_m:
            disp_act = float(comp_list[0]["Disponibilidad"].replace('%',''))
            color_v = "🟢" if disp_act > 20 else "🟡" if disp_act >= 10 else "🔴"
            st.metric("Estado de Capacidad", comp_list[0]["Estado"], delta=color_v)
            st.metric("Disponibilidad Pura", comp_list[0]["Disponibilidad"])
            st.metric("Horas Netas", comp_list[0]["Total Neto"])
            if st.button("📥 Descargar Reporte Mensual (PDF)"):
                total_bruto = res_ind[p_sel].sum()
                h_ina = res_ind[res_ind['Tarea'].isin(TAREAS_DESCUENTO_CAPACIDAD)][p_sel].sum()
                total_neto = total_bruto - h_ina
                datos_t = [["Tarea", "Horas", "% (sobre Neto)"]]
                for _, r in res_ind.iterrows():
                    porc = f"{round((r[p_sel]/total_neto)*100, 1)}%" if total_neto > 0 and r['Tarea'] not in TAREAS_DESCUENTO_CAPACIDAD else "-"
                    datos_t.append([r['Tarea'], round(r[p_sel], 1), porc])
                datos_t.append(["TOTAL CARGADO", round(total_bruto, 1), ""])
                datos_t.append(["TOTAL NETO PRODUCTIVO", round(total_neto, 1), "100%"])
                pdf_m = generar_pdf_base(f"Reporte Mensual: {p_sel}", f"{MESES_ES[mes]} {anio}[cite: 1]", [("Detalle de Jornada", datos_t)], incluir_grafico=res_ind.set_index('Tarea')[p_sel].to_dict())
                st.download_button("Guardar Mensual", pdf_m, f"Mensual_{p_sel}.pdf")

    if st.session_state.usuario_actual == "Admin - Ver todo":
        st.divider(); st.subheader("🌐 Visión Global del Estudio (Horas Netas)")
        df_eq_neta = df_act[~df_act['Tarea'].isin(TAREAS_DESCUENTO_CAPACIDAD)].melt(id_vars=['Fecha', 'Tarea'], value_vars=OPERARIOS_FIJOS, var_name='Op', value_name='Hs')
        res_eq_neta = df_eq_neta.groupby('Tarea')['Hs'].sum().reset_index()
        st.plotly_chart(px.pie(res_eq_neta, values='Hs', names='Tarea', color='Tarea', color_discrete_map=COLORES_TAREAS, title="Total Horas Productivas Equipo"), use_container_width=True)
        hist_g = {}
        for i in range(3):
            m_c = mes - i; a_c = anio
            if m_c <= 0: m_c += 12; a_c -= 1
            df_m_g = df_p[(df_p['Fecha'].dt.month == m_c) & (df_p['Fecha'].dt.year == a_c)]
            hist_g[MESES_ES[m_c]] = df_m_g[~df_m_g['Tarea'].isin(TAREAS_DESCUENTO_CAPACIDAD)][OPERARIOS_FIJOS].sum(axis=1).groupby(df_m_g['Tarea']).sum().to_dict()
        tareas_g = sorted(list(set([t for m in hist_g for t in hist_g[m].keys()]))); meses_g = list(hist_g.keys())
        rows_g = []; totales_g = [0.0] * len(meses_g)
        for t in tareas_g:
            fila = [t]
            for idx, m in enumerate(meses_g):
                val = round(float(hist_g[m].get(t, 0)), 1); fila.append(val); totales_g[idx] += val
            rows_g.append(fila)
        rows_g.append(["TOTAL NETO"] + [round(x, 1) for x in totales_g])
        st.table(pd.DataFrame(rows_g, columns=["Tarea"] + meses_g))
        if st.button("📥 Descargar Reporte Global Trimestral (PDF)"):
            pdf_g = generar_pdf_base("REPORTE GLOBAL NETO", "Estudio Completo - Grupo Pressacco[cite: 1]", [("Consolidado Productivo", [["Tarea"] + meses_g] + rows_g)], incluir_grafico=res_eq_neta.set_index('Tarea')['Hs'].to_dict())
            st.download_button("Guardar Reporte Global", pdf_g, "Global_Neto.pdf")

elif "Cargar" in menu:
    st.title("➕ Registro de Horas")
    u_c = st.session_state.usuario_actual if st.session_state.usuario_actual != "Admin - Ver todo" else st.selectbox("Persona:", OPERARIOS_FIJOS)
    with st.form("f_ind"):
        f_f = st.date_input("Fecha", value=datetime.now()); f_t = st.selectbox("Tarea", list(COLORES_TAREAS.keys())); f_h = st.number_input("Horas", step=0.5, value=6.0); f_n = st.text_input("Nota")
        if st.form_submit_button("Guardar"):
            nueva = {'Fecha': f_f, 'Tarea': f_t, 'Nota': f_n}
            for op in OPERARIOS_FIJOS: nueva[op] = f_h if op == u_c else 0
            st.session_state.cargas = pd.concat([st.session_state.cargas, pd.DataFrame([nueva])], ignore_index=True)
            guardar_df("Cargas", st.session_state.cargas); st.success("¡Guardado!"); time.sleep(1); st.rerun()
    st.divider(); st.subheader("📋 Historial y Consulta")
    mes_filt = st.selectbox("Consultar Mes:", list(range(1,13)), format_func=lambda x: MESES_ES[x], index=datetime.now().month-1)
    df_h = st.session_state.cargas.copy(); df_h['Fecha'] = pd.to_datetime(df_h['Fecha'], errors='coerce')
    df_f = df_h[(df_h[u_c] > 0) & (df_h['Fecha'].dt.month == mes_filt)]
    if not df_f.empty:
        total_bruto = df_f[u_c].sum()
        inasistencias = df_f[df_f['Tarea'].isin(TAREAS_DESCUENTO_CAPACIDAD)][u_c].sum()
        st.info(f"**Resumen {MESES_ES[mes_filt]}:** Total Cargado: {round(total_bruto, 1)} hs | Inasistencias: {round(inasistencias, 1)} hs | **Neto: {round(total_bruto - inasistencias, 1)} hs**[cite: 1]")
        st.dataframe(df_f.groupby('Tarea')[u_c].sum().round(1).reset_index(), use_container_width=True, hide_index=True)
        for i, row in df_f.sort_values('Fecha', ascending=False).iterrows():
            c1, c2 = st.columns([6, 1])
            c1.write(f"📅 {row['Fecha'].strftime('%d/%m/%Y')} | {row['Tarea']} | {row[u_c]} hs | {row['Nota']}")
            if c2.button("Eliminar", key=f"del_{i}"):
                st.session_state.cargas = st.session_state.cargas.drop(i).reset_index(drop=True); guardar_df("Cargas", st.session_state.cargas); st.rerun()

elif "Protocolo" in menu:
    st.title("📜 Protocolo de Uso - Grupo Pressacco")
    st.info("Elegimos sumar[cite: 2]")
    st.markdown("""
    ### 1. Finalidad y Objetivo
    El objetivo principal es transformar nuestra carga de trabajo en **datos accionables**. Registrar el tiempo permite medir rentabilidad, detectar saturaciones y planificar el crecimiento.[cite: 2]
    
    ### 2. Paso a Paso: Carga y Control
    *   **Registro Diario:** Cada integrante debe registrar **6 horas diarias** antes de las **15:00 hs**.
    *   **Inasistencias:** Cargarlas como tal para descontar capacidad neta automáticamente.
    *   **Autocontrol:** Verificar que el total sume 6 hs y las notas sean descriptivas.[cite: 2]
    """)
    if st.button("📥 Descargar Guía Maestra (PDF)"):
        pdf = generar_pdf_base("PROTOCOLO DE USO - CRM", "Manual de Procedimientos y Guía Completa[cite: 2]", [], es_protocolo=True)
        st.download_button("Guardar Protocolo Maestro", pdf, "Protocolo_Pressacco.pdf")

elif "Carga Masiva" in menu:
    st.title("📁 Reparto de Horas (Admin)")
    with st.form("f_m"):
        u_m = st.selectbox("Operario", OPERARIOS_FIJOS); t_m = st.selectbox("Tarea", list(COLORES_TAREAS.keys())); f_i = st.date_input("Desde"); f_f = st.date_input("Hasta"); h_t = st.number_input("Horas Totales", min_value=0.0)
        if st.form_submit_button("Distribuir"):
            dias = pd.bdate_range(start=f_i, end=f_f, freq='C', holidays=FERIADOS)
            if len(dias) > 0:
                h_d = round(h_t / len(dias), 2); filas = []
                for d in dias:
                    f = {'Fecha': d, 'Tarea': t_m, 'Nota': 'Carga Masiva'}
                    for o in OPERARIOS_FIJOS: f[o] = h_d if o == u_m else 0
                    filas.append(f)
                st.session_state.cargas = pd.concat([st.session_state.cargas, pd.DataFrame(filas)], ignore_index=True); guardar_df("Cargas", st.session_state.cargas); st.rerun()

elif "Reset" in menu:
    if st.text_input("Escriba BORRAR") == "BORRAR":
        if st.button("Eliminar Todo"): guardar_df("Cargas", pd.DataFrame(columns=['Fecha', 'Tarea'] + OPERARIOS_FIJOS + ['Nota'])); st.rerun()
