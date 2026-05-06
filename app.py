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
st.set_page_config(page_title="CRM Grupo Pressacco", layout="wide", page_icon="🏛️")
 
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
 
# ===== 2. CSS PERSONALIZADO =====
st.markdown("""
<style>
    .kpi-card {
        background: linear-gradient(135deg, #0077b6, #00b4d8);
        border-radius: 12px; padding: 18px 22px; color: white;
        text-align: center; margin-bottom: 8px;
    }
    .kpi-card h2 { font-size: 2rem; margin: 0; font-weight: 700; }
    .kpi-card p  { font-size: 0.82rem; margin: 0; opacity: 0.88; }
    .alerta-box {
        border-left: 5px solid #e63946; background: #fff0f0;
        border-radius: 6px; padding: 10px 15px; margin-bottom: 8px;
        font-size: 0.9rem; color: #333;
    }
    .desvio-pos { color: #2d6a4f; font-weight: 600; }
    .desvio-neg { color: #e63946; font-weight: 600; }
</style>
""", unsafe_allow_html=True)
 
# ===== 3. HELPERS DE FECHA Y CAPACIDAD =====
def get_rango_mes(anio, mes):
    ini = datetime(anio, mes, 1).date()
    fin = (datetime(anio, mes + 1, 1) if mes < 12 else datetime(anio + 1, 1, 1)).date() - timedelta(days=1)
    return ini, fin
 
def capacidad_neta(anio, mes, feriados, horas_ina=0):
    ini, fin = get_rango_mes(anio, mes)
    dias_h = len(pd.bdate_range(start=ini, end=fin, freq='C', holidays=feriados))
    return round(dias_h * HORAS_DIA_LABORAL - horas_ina, 1)
 
def semaforo_estado(disp_pct):
    if disp_pct > 20: return "🟢 Libre"
    elif disp_pct >= 10: return "🟡 Atención"
    else: return "🔴 Saturado"
 
# ===== 4. ANÁLISIS DE DESVÍO VS HISTÓRICO =====
def calcular_desvio(df, operario, mes_actual, anio_actual):
    """Compara horas por tarea del mes actual vs promedio de los 2 meses anteriores."""
    meses_hist = []
    for i in range(1, 3):
        m = mes_actual - i; a = anio_actual
        if m <= 0: m += 12; a -= 1
        meses_hist.append((a, m))
 
    df_act = df[(df['Fecha'].dt.month == mes_actual) & (df['Fecha'].dt.year == anio_actual)]
    act_por_tarea = df_act.groupby('Tarea')[operario].sum()
 
    hist_frames = []
    for a, m in meses_hist:
        df_m = df[(df['Fecha'].dt.month == m) & (df['Fecha'].dt.year == a)]
        hist_frames.append(df_m.groupby('Tarea')[operario].sum())
 
    if not hist_frames:
        return pd.DataFrame()
 
    prom_hist = pd.concat(hist_frames, axis=1).fillna(0).mean(axis=1)
    desvio = pd.DataFrame({
        'Tarea': act_por_tarea.index.union(prom_hist.index),
    })
    desvio = desvio.set_index('Tarea')
    desvio['Actual (hs)'] = act_por_tarea.reindex(desvio.index).fillna(0).round(1)
    desvio['Promedio Hist. (hs)'] = prom_hist.reindex(desvio.index).fillna(0).round(1)
    desvio['Desvío (hs)'] = (desvio['Actual (hs)'] - desvio['Promedio Hist. (hs)']).round(1)
    desvio['Desvío (%)'] = desvio.apply(
        lambda r: round((r['Desvío (hs)'] / r['Promedio Hist. (hs)']) * 100, 1) if r['Promedio Hist. (hs)'] > 0 else 0, axis=1
    )
    return desvio[desvio['Actual (hs)'] > 0].sort_values('Desvío (hs)', ascending=False).reset_index()
 
# ===== 5. TENDENCIA HISTÓRICA (6 MESES) =====
def tendencia_historica(df, operario, mes_actual, anio_actual):
    """Retorna DataFrame con horas por tarea para los últimos 6 meses."""
    rows = []
    for i in range(5, -1, -1):
        m = mes_actual - i; a = anio_actual
        if m <= 0: m += 12; a -= 1
        df_m = df[(df['Fecha'].dt.month == m) & (df['Fecha'].dt.year == a)]
        por_tarea = df_m.groupby('Tarea')[operario].sum()
        for tarea, hs in por_tarea.items():
            if hs > 0:
                rows.append({'Mes': MESES_ES[m], 'Orden': i, 'Tarea': tarea, 'Horas': round(hs, 1)})
    return pd.DataFrame(rows)
 
# ===== 6. DISTRIBUCIÓN SEMANAL =====
def distribucion_semanal(df, operario, anio, mes):
    df_m = df[(df['Fecha'].dt.month == mes) & (df['Fecha'].dt.year == anio)].copy()
    if df_m.empty: return pd.DataFrame()
    df_m['Semana'] = df_m['Fecha'].dt.isocalendar().week.astype(str).apply(lambda w: f"Sem {w}")
    return df_m.groupby(['Semana', 'Tarea'])[operario].sum().reset_index()
 
# ===== 7. BALANCE DEL EQUIPO =====
def balance_equipo(df, anio, mes, feriados):
    df_m = df[(df['Fecha'].dt.month == mes) & (df['Fecha'].dt.year == anio)].copy()
    rows = []
    for op in OPERARIOS_FIJOS:
        h_ina = df_m[df_m['Tarea'].isin(TAREAS_DESCUENTO_CAPACIDAD)][op].sum()
        cap = capacidad_neta(anio, mes, feriados, h_ina)
        total = round(df_m[op].sum(), 1)
        h_disp = df_m[df_m['Tarea'].isin(TAREAS_DISPONIBILIDAD)][op].sum()
        disp_pct = round((h_disp / cap * 100) if cap > 0 else 0, 1)
        util_pct = round(((total - h_disp) / cap * 100) if cap > 0 else 0, 1)
        rows.append({
            'Integrante': op, 'Horas Cargadas': total, 'Capacidad Neta': cap,
            'Disponibilidad %': disp_pct, 'Utilización %': util_pct,
            'Estado': semaforo_estado(disp_pct)
        })
    return pd.DataFrame(rows)
 
# ===== 8. ALERTAS AUTOMÁTICAS =====
def generar_alertas(df, operario, anio, mes, feriados):
    alertas = []
    df_m = df[(df['Fecha'].dt.month == mes) & (df['Fecha'].dt.year == anio)].copy()
    ini, fin = get_rango_mes(anio, mes)
    dias_lab = pd.bdate_range(start=ini, end=datetime.now().date(), freq='C', holidays=feriados)
 
    # Días laborales sin carga
    fechas_cargadas = set(df_m[df_m[operario] > 0]['Fecha'].dt.date)
    dias_sin_carga = [d.date() for d in dias_lab if d.date() not in fechas_cargadas]
    if dias_sin_carga:
        alertas.append(f"📅 {len(dias_sin_carga)} día(s) sin carga: {', '.join(str(d.strftime('%d/%m')) for d in dias_sin_carga[-3:])}" +
                       (" y más..." if len(dias_sin_carga) > 3 else ""))
 
    # Desvío fuerte (>50%) en alguna tarea
    dev = calcular_desvio(df, operario, mes, anio)
    if not dev.empty:
        grandes = dev[abs(dev['Desvío (%)']) > 50]
        for _, row in grandes.iterrows():
            signo = "+" if row['Desvío (hs)'] > 0 else ""
            alertas.append(f"⚠️ {row['Tarea']}: desvío de {signo}{row['Desvío (hs)']} hs ({signo}{row['Desvío (%)']}%) vs. promedio histórico")
 
    # Saturación (disponibilidad < 10%)
    df_disp = df_m[df_m['Tarea'].isin(TAREAS_DISPONIBILIDAD)][operario].sum()
    h_ina = df_m[df_m['Tarea'].isin(TAREAS_DESCUENTO_CAPACIDAD)][operario].sum()
    cap = capacidad_neta(anio, mes, feriados, h_ina)
    disp_pct = (df_disp / cap * 100) if cap > 0 else 0
    if disp_pct < 10 and cap > 0:
        alertas.append(f"🔴 Saturación crítica: solo {round(disp_pct, 1)}% de disponibilidad este mes")
 
    return alertas
 
# ===== 9. PDF MEJORADO =====
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
    estilo_negrita = s['Normal']; estilo_negrita.fontSize = 9; estilo_negrita.fontName = 'Helvetica-Bold'
 
    story = [Paragraph("GRUPO PRESSACCO", estilo_titulo), Paragraph(titulo_doc, s['Heading2']),
             Paragraph(subtitulo, s['Normal']), Spacer(1, 15)]
 
    if es_protocolo:
        for t, c in [("1. FINALIDAD", "Transformar carga de trabajo en datos accionables."),
                     ("2. OBJETIVOS", "• Visibilidad. • Equilibrio. • Transparencia."),
                     ("3. REGLAS", "• Carga 6hs diarias antes de las 15hs. • Inasistencias descuentan capacidad.")]:
            story.append(Paragraph(t, s['Heading3'])); story.append(Paragraph(c, estilo_cuerpo)); story.append(Spacer(1, 10))
 
    if incluir_grafico:
        d = Drawing(450, 200); pc = Pie(); pc.x = 50; pc.y = 25; pc.width = 130; pc.height = 130
        lista_colores = [colors.magenta, colors.deepskyblue, colors.lightpink, colors.yellow,
                         colors.whitesmoke, colors.lightblue, colors.lavender, colors.bisque,
                         colors.turquoise, colors.lime, colors.hotpink]
        grafico_limpio = {k: v for k, v in incluir_grafico.items() if k not in TAREAS_DESCUENTO_CAPACIDAD and v > 0}
        total_h = sum(grafico_limpio.values())
        if total_h > 0:
            pc.data = [round(float(v), 1) for v in grafico_limpio.values()]
            pc.labels = [f"{round((v/total_h)*100, 1)}%" for v in grafico_limpio.values()]
            for i in range(len(pc.data)): pc.slices[i].fillColor = lista_colores[i % len(lista_colores)]
            leg = Legend(); leg.x = 220; leg.y = 150; leg.alignment = 'right'
            leg.columnMaximum = 12; leg.fontSize = 7
            leg.colorNamePairs = [(lista_colores[i % len(lista_colores)], k) for i, k in enumerate(grafico_limpio.keys())]
            d.add(pc); d.add(leg); story.append(d); story.append(Spacer(1, 10))
 
    for t_tabla, data in datos_tablas:
        if t_tabla: story.append(Paragraph(t_tabla, s['Heading3']))
        data_p = [[Paragraph(str(c), estilo_negrita if "TOTAL" in str(c).upper() else estilo_cuerpo) for c in fila] for fila in data]
        col_w = [2.5*inch] + [1.0*inch] * (len(data[0]) - 1)
        t = Table(data_p, colWidths=col_w)
        t.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), color_celeste),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')
        ]))
        story.append(t); story.append(Spacer(1, 15))
 
    doc.build(story); buf.seek(0); return buf
 
# ===== 10. CONEXIÓN GOOGLE SHEETS =====
@st.cache_resource
def conectar():
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    )
    return gspread.authorize(creds).open("CRM_Estudio_Datos")
 
def cargar_hoja(nombre):
    try:
        ws = conectar().worksheet(nombre)
        df = pd.DataFrame(ws.get_all_records())
        for c in df.columns:
            if 'fecha' in c.lower():
                df[c] = pd.to_datetime(df[c], dayfirst=True, errors='coerce')
        return df, ws
    except:
        return pd.DataFrame(), None
 
def guardar_df(nombre, df):
    try:
        ws = conectar().worksheet(nombre); ws.clear()
        df_c = df.copy()
        for c in df_c.columns:
            if 'fecha' in c.lower():
                df_c[c] = pd.to_datetime(df_c[c], errors='coerce').dt.strftime('%d/%m/%Y')
        ws.update([df_c.columns.values.tolist()] + df_c.fillna('').astype(str).values.tolist())
        st.cache_data.clear(); return True
    except:
        return False
 
@st.cache_data
def cargar_feriados():
    try:
        df_f = pd.read_csv("feriados_2026.csv"); df_f.columns = df_f.columns.str.strip().str.lower()
        return pd.to_datetime(df_f['fecha'], errors='coerce').dt.date.dropna().tolist()
    except:
        return []
 
FERIADOS = cargar_feriados()
 
# ===== 11. ESTADO INICIAL =====
if 'cargas' not in st.session_state:
    df, _ = cargar_hoja("Cargas")
    st.session_state.cargas = df if not df.empty else pd.DataFrame(columns=['Fecha', 'Tarea'] + OPERARIOS_FIJOS + ['Nota'])
 
if 'usuario_actual' not in st.session_state:
    st.session_state.usuario_actual = None
 
# ===== 12. LOGIN =====
if st.session_state.usuario_actual is None:
    st.title("🏛️ CRM Grupo Pressacco")
    u = st.selectbox("Usuario:", ["Seleccionar..."] + OPERARIOS_FIJOS + ["Admin - Ver todo"])
    if st.button("Ingresar") and u != "Seleccionar...":
        st.session_state.usuario_actual = u; st.rerun()
    st.stop()
 
# ===== 13. ALERTA MENSUAL =====
hoy = datetime.now()
df_u = st.session_state.cargas.copy()
df_u['Fecha'] = pd.to_datetime(df_u['Fecha'], errors='coerce')
ini_m, fin_m = get_rango_mes(hoy.year, hoy.month)
dias_h = len(pd.bdate_range(start=ini_m, end=fin_m, freq='C', holidays=FERIADOS))
total_obj = dias_h * HORAS_DIA_LABORAL
 
if st.session_state.usuario_actual != "Admin - Ver todo":
    cargadas = df_u[
        (df_u['Fecha'].dt.month == hoy.month) &
        (df_u['Fecha'].dt.year == hoy.year)
    ][st.session_state.usuario_actual].sum()
    restante = total_obj - cargadas
    if restante > 0:
        st.warning(f"🎯 **Objetivo {MESES_ES[hoy.month]}:** Faltan **{round(restante, 1)} hs** para las {total_obj} hs del mes.")
    else:
        st.success(f"✅ ¡Objetivo de {total_obj} hs cumplido!")
 
# ===== 14. NAVEGACIÓN =====
es_admin = st.session_state.usuario_actual == "Admin - Ver todo"
menu_opciones = ["📊 Panel de Control", "➕ Cargar Horas", "📁 Carga Masiva", "📜 Protocolo", "⚙️ Reset"] if es_admin \
    else ["📊 Panel de Control", "➕ Cargar Mis Horas", "📜 Protocolo"]
menu = st.sidebar.radio("Navegación", menu_opciones)
if st.sidebar.button("Cerrar Sesión"):
    st.session_state.clear(); st.rerun()
 
# ===== 15. PANEL DE CONTROL =====
if "Panel de Control" in menu:
    st.title("📊 Análisis de Capacidad y Eficiencia")
 
    c1, c2, c3 = st.columns([1, 1, 2])
    with c1: anio = st.selectbox("Año", [2025, 2026], index=1)
    with c2: mes = st.selectbox("Mes", list(range(1, 13)), format_func=lambda x: MESES_ES[x], index=hoy.month - 1)
    with c3:
        p_sel = st.selectbox("Integrante:", OPERARIOS_FIJOS) if es_admin else st.session_state.usuario_actual
 
    df_p = st.session_state.cargas.copy()
    df_p['Fecha'] = pd.to_datetime(df_p['Fecha'], errors='coerce')
    df_act = df_p[(df_p['Fecha'].dt.month == mes) & (df_p['Fecha'].dt.year == anio)]
 
    # --- KPIs ---
    h_ina = df_act[df_act['Tarea'].isin(TAREAS_DESCUENTO_CAPACIDAD)][p_sel].sum()
    cap = capacidad_neta(anio, mes, FERIADOS, h_ina)
    total_cargado = round(df_act[p_sel].sum(), 1)
    h_disp = df_act[df_act['Tarea'].isin(TAREAS_DISPONIBILIDAD)][p_sel].sum()
    disp_pct = round((h_disp / cap * 100) if cap > 0 else 0, 1)
    util_pct = round(((total_cargado - h_disp) / cap * 100) if cap > 0 else 0, 1)
 
    k1, k2, k3, k4 = st.columns(4)
    with k1: st.markdown(f'<div class="kpi-card"><h2>{total_cargado}</h2><p>Horas Cargadas</p></div>', unsafe_allow_html=True)
    with k2: st.markdown(f'<div class="kpi-card"><h2>{cap}</h2><p>Capacidad Neta</p></div>', unsafe_allow_html=True)
    with k3: st.markdown(f'<div class="kpi-card"><h2>{util_pct}%</h2><p>Utilización Real</p></div>', unsafe_allow_html=True)
    with k4: st.markdown(f'<div class="kpi-card"><h2>{disp_pct}%</h2><p>Disponibilidad {semaforo_estado(disp_pct)}</p></div>', unsafe_allow_html=True)
 
    st.divider()
 
    # --- ALERTAS AUTOMÁTICAS ---
    alertas = generar_alertas(df_p, p_sel, anio, mes, FERIADOS)
    if alertas:
        st.subheader("🚨 Alertas Automáticas")
        for a in alertas:
            st.markdown(f'<div class="alerta-box">{a}</div>', unsafe_allow_html=True)
        st.divider()
 
    # --- ANÁLISIS DE DESVÍOS ---
    st.subheader(f"📐 Desvío vs. Promedio Histórico — {p_sel}")
    dev = calcular_desvio(df_p, p_sel, mes, anio)
    if not dev.empty:
        col_t, col_g = st.columns([1, 1])
        with col_t:
            def colorear(val):
                if isinstance(val, (int, float)):
                    if val > 0: return 'color: #2d6a4f; font-weight:600'
                    elif val < 0: return 'color: #e63946; font-weight:600'
                return ''
            st.dataframe(
                dev.style.applymap(colorear, subset=['Desvío (hs)', 'Desvío (%)']),
                use_container_width=True, hide_index=True
            )
        with col_g:
            fig_dev = px.bar(dev, x='Desvío (hs)', y='Tarea', orientation='h',
                             color='Desvío (hs)', color_continuous_scale=['#e63946', '#adb5bd', '#2d6a4f'],
                             title="Desvío en horas (actual vs. prom. histórico)")
            fig_dev.update_layout(coloraxis_showscale=False, height=350, margin=dict(l=0, r=0, t=40, b=0))
            st.plotly_chart(fig_dev, use_container_width=True)
    else:
        st.info("No hay datos históricos suficientes para calcular desvíos (se necesitan al menos 2 meses previos).")
 
    st.divider()
 
    # --- TENDENCIA HISTÓRICA 6 MESES ---
    st.subheader(f"📈 Tendencia Histórica — {p_sel} (últimos 6 meses)")
    df_tend = tendencia_historica(df_p, p_sel, mes, anio)
    if not df_tend.empty:
        orden_meses = df_tend.sort_values('Orden', ascending=False)['Mes'].unique().tolist()
        fig_tend = px.bar(df_tend, x='Mes', y='Horas', color='Tarea',
                          color_discrete_map=COLORES_TAREAS,
                          category_orders={'Mes': orden_meses},
                          title="Horas por tarea — evolución mensual")
        fig_tend.update_layout(barmode='stack', height=380, margin=dict(l=0, r=0, t=40, b=0))
        st.plotly_chart(fig_tend, use_container_width=True)
 
    st.divider()
 
    # --- DISTRIBUCIÓN SEMANAL ---
    st.subheader(f"📅 Distribución Semanal — {MESES_ES[mes]} {anio}")
    df_sem = distribucion_semanal(df_p, p_sel, anio, mes)
    if not df_sem.empty:
        fig_sem = px.bar(df_sem, x='Semana', y=p_sel, color='Tarea',
                         color_discrete_map=COLORES_TAREAS, barmode='stack',
                         title="Carga por semana")
        fig_sem.update_layout(height=320, margin=dict(l=0, r=0, t=40, b=0))
        st.plotly_chart(fig_sem, use_container_width=True)
    else:
        st.info("Sin datos para distribución semanal.")
 
    st.divider()
 
    # --- DONA INDIVIDUAL ---
    res_ind = df_act.groupby('Tarea')[p_sel].sum().round(1).reset_index()
    res_graf = res_ind[(res_ind[p_sel] > 0) & (~res_ind['Tarea'].isin(TAREAS_DESCUENTO_CAPACIDAD))]
    if not res_graf.empty:
        st.subheader(f"🍩 Composición de Horas — {p_sel}")
        col_g, col_m = st.columns([2, 1])
        with col_g:
            fig = px.pie(res_graf, values=p_sel, names='Tarea', color='Tarea',
                         color_discrete_map=COLORES_TAREAS, hole=0.5,
                         title=f"Eficiencia Real — {p_sel}")
            fig.update_traces(textposition='inside', textinfo='percent+label')
            st.plotly_chart(fig, use_container_width=True)
        with col_m:
            if st.button("📥 PDF Mensual Individual"):
                dat = [["Tarea", "Horas"]] + [[r['Tarea'], r[p_sel]] for _, r in res_ind.iterrows()] + [["TOTAL", total_cargado]]
                st.download_button("Guardar Mensual",
                    generar_pdf_base(f"Reporte {p_sel}", f"{MESES_ES[mes]} {anio}", [("Detalle", dat)],
                                     incluir_grafico=res_ind.set_index('Tarea')[p_sel].to_dict()),
                    f"Mensual_{p_sel}.pdf")
 
    # --- COMPARATIVA TRIMESTRAL ---
    st.divider()
    st.subheader(f"🗂️ Comparativa Trimestral — {p_sel}")
    comp_list = []; hist_pdf = {}
    for i in range(3):
        m_c = mes - i; a_c = anio
        if m_c <= 0: m_c += 12; a_c -= 1
        df_m = df_p[(df_p['Fecha'].dt.month == m_c) & (df_p['Fecha'].dt.year == a_c)]
        h_ina_c = df_m[df_m['Tarea'].isin(TAREAS_DESCUENTO_CAPACIDAD)][p_sel].sum()
        cap_n = capacidad_neta(a_c, m_c, FERIADOS, h_ina_c)
        total_b = round(df_m[p_sel].sum(), 1)
        h_disp_c = df_m[df_m['Tarea'].isin(TAREAS_DISPONIBILIDAD)][p_sel].sum()
        disp_v = round((h_disp_c / cap_n * 100) if cap_n > 0 else 0, 1)
        comp_list.append({"Mes": MESES_ES[m_c], "Carga": f"{total_b} hs", "Cap. Neta": f"{cap_n} hs",
                           "Disponibilidad": f"{disp_v}%", "Estado": semaforo_estado(disp_v)})
        hist_pdf[MESES_ES[m_c]] = df_m.groupby('Tarea')[p_sel].sum().to_dict()
 
    st.table(pd.DataFrame(comp_list))
    if st.button("📥 PDF Trimestral Individual"):
        tareas_u = sorted(list(set([t for m in hist_pdf for t in hist_pdf[m].keys()])))
        meses_n = list(hist_pdf.keys()); rows = []; tot_m = [0.0] * len(meses_n)
        for t in tareas_u:
            f = [t]
            for idx, m in enumerate(meses_n):
                v = round(float(hist_pdf[m].get(t, 0)), 1); f.append(v); tot_m[idx] += v
            rows.append(f)
        rows.append(["TOTAL BRUTO"] + [round(x, 1) for x in tot_m])
        st.download_button("Guardar Trimestral",
            generar_pdf_base(f"Trimestral: {p_sel}", "Resumen 3 meses",
                             [("Detalle", [["Tarea"] + meses_n] + rows)]),
            f"Trimestral_{p_sel}.pdf")
 
    # --- VISIÓN GLOBAL ADMIN ---
    if es_admin:
        st.divider(); st.subheader("🌐 Balance del Equipo")
        df_bal = balance_equipo(df_p, anio, mes, FERIADOS)
        st.dataframe(df_bal, use_container_width=True, hide_index=True)
 
        # Heatmap de carga por integrante
        df_heat = df_act[~df_act['Tarea'].isin(TAREAS_DESCUENTO_CAPACIDAD)].copy()
        if not df_heat.empty:
            heat_data = df_heat.groupby('Tarea')[OPERARIOS_FIJOS].sum().round(1)
            fig_heat = go.Figure(data=go.Heatmap(
                z=heat_data.values, x=heat_data.columns.tolist(), y=heat_data.index.tolist(),
                colorscale='Blues', text=heat_data.values, texttemplate="%{text}",
                showscale=True
            ))
            fig_heat.update_layout(title="Heatmap de Horas por Tarea e Integrante",
                                   height=400, margin=dict(l=0, r=0, t=40, b=0))
            st.plotly_chart(fig_heat, use_container_width=True)
 
        # Pie global
        df_eq = df_act[~df_act['Tarea'].isin(TAREAS_DESCUENTO_CAPACIDAD)].melt(
            id_vars=['Fecha', 'Tarea'], value_vars=OPERARIOS_FIJOS, var_name='Op', value_name='Hs')
        res_eq = df_eq.groupby('Tarea')['Hs'].sum().reset_index()
        fig_g = px.pie(res_eq, values='Hs', names='Tarea', color='Tarea',
                       color_discrete_map=COLORES_TAREAS, hole=0.5, title="Total Horas Equipo (Netas)")
        fig_g.update_traces(textposition='inside', textinfo='percent+label')
        st.plotly_chart(fig_g, use_container_width=True)
 
        if st.button("📥 PDF Global"):
            hist_g = {}
            for i in range(3):
                m_c = mes - i; a_c = anio
                if m_c <= 0: m_c += 12; a_c -= 1
                df_m_g = df_p[(df_p['Fecha'].dt.month == m_c) & (df_p['Fecha'].dt.year == a_c)]
                vals = df_m_g[~df_m_g['Tarea'].isin(TAREAS_DESCUENTO_CAPACIDAD)][OPERARIOS_FIJOS].sum(axis=1)
                hist_g[MESES_ES[m_c]] = vals.groupby(df_m_g['Tarea']).sum().to_dict()
            tg = sorted(list(set([t for m in hist_g for t in hist_g[m].keys()])))
            mg = list(hist_g.keys()); rg = []; totg = [0.0] * len(mg)
            for t in tg:
                f = [t]
                for idx, m in enumerate(mg):
                    v = round(float(hist_g[m].get(t, 0)), 1); f.append(v); totg[idx] += v
                rg.append(f)
            rg.append(["TOTAL NETO"] + [round(x, 1) for x in totg])
            st.download_button("Guardar Global",
                generar_pdf_base("REPORTE GLOBAL", "Estudio Completo",
                                 [("Consolidado", [["Tarea"] + mg] + rg)],
                                 incluir_grafico=res_eq.set_index('Tarea')['Hs'].to_dict()),
                "Global_Neto.pdf")
 
# ===== 16. CARGA MASIVA =====
elif "Carga Masiva" in menu:
    st.title("📁 Distribución Masiva")
    with st.form("fm"):
        u_m = st.selectbox("Persona", OPERARIOS_FIJOS)
        t_m = st.selectbox("Tarea", list(COLORES_TAREAS.keys()))
        f_i = st.date_input("Desde"); f_f = st.date_input("Hasta")
        h_t = st.number_input("Horas Totales", min_value=0.0)
        if st.form_submit_button("Ejecutar"):
            rgo = pd.bdate_range(start=f_i, end=f_f, freq='C', holidays=FERIADOS)
            if len(rgo) > 0:
                h_d = round(h_t / len(rgo), 2); filas = []
                for d in rgo:
                    fila = {'Fecha': d, 'Tarea': t_m, 'Nota': 'Masiva'}
                    for o in OPERARIOS_FIJOS: fila[o] = h_d if o == u_m else 0
                    filas.append(fila)
                st.session_state.cargas = pd.concat([st.session_state.cargas, pd.DataFrame(filas)], ignore_index=True)
                if guardar_df("Cargas", st.session_state.cargas):
                    st.success("✅ Guardado."); time.sleep(1.5); st.rerun()
 
# ===== 17. CARGA INDIVIDUAL =====
elif "Cargar" in menu:
    st.title("➕ Registro de Horas")
    u_c = st.session_state.usuario_actual if not es_admin else st.selectbox("Persona:", OPERARIOS_FIJOS)
 
    with st.form("fi"):
        f_fecha = st.date_input("Fecha", value=datetime.now())
        f_t = st.selectbox("Tarea", list(COLORES_TAREAS.keys()))
        f_h = st.number_input("Horas", step=0.5, value=6.0)
        f_n = st.text_input("Nota")
        if st.form_submit_button("Guardar"):
            df_check = st.session_state.cargas.copy()
            df_check['Fecha'] = pd.to_datetime(df_check['Fecha'], errors='coerce')
            # Validación de duplicado
            duplicado = df_check[
                (df_check['Fecha'].dt.date == f_fecha) &
                (df_check['Tarea'] == f_t) &
                (df_check[u_c] > 0)
            ]
            if not duplicado.empty:
                st.warning(f"⚠️ Ya existe una carga para {u_c} en {f_fecha.strftime('%d/%m/%Y')} con tarea '{f_t}'. Eliminá la anterior si querés reemplazarla.")
            elif f_h > HORAS_DIA_LABORAL:
                st.warning(f"⚠️ Estás cargando {f_h} hs para un día (máximo habitual: {HORAS_DIA_LABORAL} hs). ¿Seguro?")
                nueva = {'Fecha': f_fecha, 'Tarea': f_t, 'Nota': f_n}
                for op in OPERARIOS_FIJOS: nueva[op] = f_h if op == u_c else 0
                st.session_state.cargas = pd.concat([st.session_state.cargas, pd.DataFrame([nueva])], ignore_index=True)
                guardar_df("Cargas", st.session_state.cargas); st.success("¡Guardado con advertencia!"); time.sleep(1); st.rerun()
            else:
                nueva = {'Fecha': f_fecha, 'Tarea': f_t, 'Nota': f_n}
                for op in OPERARIOS_FIJOS: nueva[op] = f_h if op == u_c else 0
                st.session_state.cargas = pd.concat([st.session_state.cargas, pd.DataFrame([nueva])], ignore_index=True)
                guardar_df("Cargas", st.session_state.cargas); st.success("¡Guardado!"); time.sleep(1); st.rerun()
 
    st.divider()
    mes_f = st.selectbox("Historial Mes:", list(range(1, 13)), format_func=lambda x: MESES_ES[x], index=hoy.month - 1)
    df_h = st.session_state.cargas.copy(); df_h['Fecha'] = pd.to_datetime(df_h['Fecha'], errors='coerce')
    df_f = df_h[(df_h[u_c] > 0) & (df_h['Fecha'].dt.month == mes_f)]
    if not df_f.empty:
        st.info(f"**Total {MESES_ES[mes_f]}:** {round(df_f[u_c].sum(), 1)} hs")
        for i, r in df_f.sort_values('Fecha', ascending=False).iterrows():
            c1, c2 = st.columns([6, 1])
            c1.write(f"📅 {r['Fecha'].strftime('%d/%m/%Y')} | {r['Tarea']} | {r[u_c]} hs | {r['Nota']}")
            if c2.button("Eliminar", key=f"del_{i}"):
                st.session_state.cargas.drop(i, inplace=True)
                guardar_df("Cargas", st.session_state.cargas); st.rerun()
 
# ===== 18. PROTOCOLO =====
elif "Protocolo" in menu:
    st.title("📜 Protocolo de Uso")
    if st.button("📥 Descargar Guía"):
        st.download_button("Guardar",
            generar_pdf_base("PROTOCOLO CRM", "Manual de Procedimientos", [], es_protocolo=True),
            "Protocolo_Pressacco.pdf")
 
# ===== 19. RESET =====
elif "Reset" in menu:
    st.title("⚙️ Resetear Base")
    if st.text_input("Escriba BORRAR") == "BORRAR":
        if st.button("Limpiar Todo"):
            guardar_df("Cargas", pd.DataFrame(columns=['Fecha', 'Tarea'] + OPERARIOS_FIJOS + ['Nota'])); st.rerun()
