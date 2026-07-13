import streamlit as st
import pandas as pd
import numpy as np
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
    "PLANIFICACIONES/ORGANIZACIÓN/PROCEDIMIENTOS/INFORMES": "#40E0D0"
}

SUBTAREAS = {
    "IMPUESTOS": ["Anual", "Mensual"],
    "SUELDOS": ["Liquidación", "Cargas Sociales y Sindicato", "SAC"],
    "CONTABILIDAD": ["Rutinaria", "Cierre y Balances"],
    "DOCUMENTACIÓN CARGA": ["Monotributo", "Responsable Inscripto", "Sociedades"],
    "DOCUMENTACIÓN CONTROL": [],
    "ATENCION AL CLIENTE": [],
    "TAREAS NO RUTINARIAS": [],
    "DISPONIBLE": [],
    "REUNIONES DE EQUIPO": [],
    "INASISTENCIA POR EXAMEN O TRAMITE": [],
    "PLANIFICACIONES/ORGANIZACIÓN/PROCEDIMIENTOS/INFORMES": [],
}

OPERARIOS_FIJOS = ["Natalia", "Maximiliano", "Athina", "Johana"]
HORAS_DIA_LABORAL = 6
HORAS_MAX_DIA = 12  # tope duro para evitar cargas absurdas (ej: 40hs en un día)
TAREAS_DESCUENTO_CAPACIDAD = ["INASISTENCIA POR EXAMEN O TRAMITE"]
TAREAS_DISPONIBILIDAD = ["DISPONIBLE", "PLANIFICACIONES/ORGANIZACIÓN/PROCEDIMIENTOS/INFORMES"]
MESES_ES = {1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
            7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"}
ANIOS_FERIADOS_A_CARGAR = [2024, 2025, 2026, 2027]  # agregar más a futuro, ver Manual

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
    .dia-card {
        border-radius: 10px; padding: 12px 16px; margin-bottom: 8px;
        border-left: 6px solid #ccc; background: #fafafa;
    }
    .dia-verde { border-left-color: #2d6a4f; background: #f0fff4; }
    .dia-amarillo { border-left-color: #e9c46a; background: #fffbea; }
    .dia-rojo { border-left-color: #e63946; background: #fff0f0; }
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

def semaforo_dia(margen_hs):
    if margen_hs <= 0: return "🔴 Saturado", "dia-rojo"
    elif margen_hs < 2: return "🟡 Ajustado", "dia-amarillo"
    else: return "🟢 Con margen", "dia-verde"

# ===== 4. ETIQUETA TAREA+SUBTAREA =====
def etiqueta_tarea(tarea, subtarea):
    if subtarea and str(subtarea).strip() and str(subtarea).strip() != "—":
        return f"{tarea} — {subtarea}"
    return tarea

# ===== 5. ANÁLISIS DE DESVÍO VS HISTÓRICO =====
def calcular_desvio(df, operario, mes_actual, anio_actual):
    meses_hist = []
    for i in range(1, 3):
        m = mes_actual - i; a = anio_actual
        if m <= 0: m += 12; a -= 1
        meses_hist.append((a, m))

    df_act = df[(df['Fecha'].dt.month == mes_actual) & (df['Fecha'].dt.year == anio_actual)].copy()
    df_act['Etiqueta'] = df_act.apply(lambda r: etiqueta_tarea(r['Tarea'], r.get('Subtarea', '')), axis=1)
    act_por_tarea = df_act.groupby('Etiqueta')[operario].sum()

    hist_frames = []
    for a, m in meses_hist:
        df_m = df[(df['Fecha'].dt.month == m) & (df['Fecha'].dt.year == a)].copy()
        df_m['Etiqueta'] = df_m.apply(lambda r: etiqueta_tarea(r['Tarea'], r.get('Subtarea', '')), axis=1)
        hist_frames.append(df_m.groupby('Etiqueta')[operario].sum())

    if not hist_frames:
        return pd.DataFrame()

    prom_hist = pd.concat(hist_frames, axis=1).fillna(0).mean(axis=1)
    desvio = pd.DataFrame({'Tarea': act_por_tarea.index.union(prom_hist.index)}).set_index('Tarea')
    desvio['Actual (hs)'] = act_por_tarea.reindex(desvio.index).fillna(0).round(1)
    desvio['Promedio Hist. (hs)'] = prom_hist.reindex(desvio.index).fillna(0).round(1)
    desvio['Desvío (hs)'] = (desvio['Actual (hs)'] - desvio['Promedio Hist. (hs)']).round(1)
    desvio['Desvío (%)'] = desvio.apply(
        lambda r: round((r['Desvío (hs)'] / r['Promedio Hist. (hs)']) * 100, 1) if r['Promedio Hist. (hs)'] > 0 else 0, axis=1
    )
    return desvio[desvio['Actual (hs)'] > 0].sort_values('Desvío (hs)', ascending=False).reset_index()

# ===== 6. TENDENCIA HISTÓRICA (6 MESES) =====
def tendencia_historica(df, operario, mes_actual, anio_actual):
    rows = []
    for i in range(5, -1, -1):
        m = mes_actual - i; a = anio_actual
        if m <= 0: m += 12; a -= 1
        df_m = df[(df['Fecha'].dt.month == m) & (df['Fecha'].dt.year == a)].copy()
        df_m['Etiqueta'] = df_m.apply(lambda r: etiqueta_tarea(r['Tarea'], r.get('Subtarea', '')), axis=1)
        por_tarea = df_m.groupby('Etiqueta')[operario].sum()
        for tarea, hs in por_tarea.items():
            if hs > 0:
                rows.append({'Mes': MESES_ES[m], 'Orden': i, 'Tarea': tarea, 'Horas': round(hs, 1)})
    return pd.DataFrame(rows)

# ===== 7. DISTRIBUCIÓN SEMANAL =====
def distribucion_semanal(df, operario, anio, mes):
    df_m = df[(df['Fecha'].dt.month == mes) & (df['Fecha'].dt.year == anio)].copy()
    if df_m.empty: return pd.DataFrame()
    df_m['Etiqueta'] = df_m.apply(lambda r: etiqueta_tarea(r['Tarea'], r.get('Subtarea', '')), axis=1)
    df_m['Semana'] = df_m['Fecha'].dt.isocalendar().week.astype(str).apply(lambda w: f"Sem {w}")
    return df_m.groupby(['Semana', 'Etiqueta'])[operario].sum().reset_index().rename(columns={'Etiqueta': 'Tarea'})

# ===== 8. BALANCE DEL EQUIPO =====
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
        dias_sat, dias_tot = dias_saturados_mes(df, op, anio, mes, feriados)
        rows.append({
            'Integrante': op, 'Horas Cargadas': total, 'Capacidad Neta': cap,
            'Disponibilidad %': disp_pct, 'Utilización %': util_pct,
            'Días Saturados': f"{dias_sat}/{dias_tot}",
            'Estado': semaforo_estado(disp_pct)
        })
    return pd.DataFrame(rows)

# ===== 9. ALERTAS AUTOMÁTICAS =====
def generar_alertas(df, operario, anio, mes, feriados):
    alertas = []
    df_m = df[(df['Fecha'].dt.month == mes) & (df['Fecha'].dt.year == anio)].copy()
    ini, fin = get_rango_mes(anio, mes)
    dias_lab = pd.bdate_range(start=ini, end=datetime.now().date(), freq='C', holidays=feriados)

    fechas_cargadas = set(df_m[df_m[operario] > 0]['Fecha'].dt.date)
    dias_sin_carga = [d.date() for d in dias_lab if d.date() not in fechas_cargadas]
    if dias_sin_carga:
        alertas.append(f"📅 {len(dias_sin_carga)} día(s) sin carga: {', '.join(str(d.strftime('%d/%m')) for d in dias_sin_carga[-3:])}" +
                       (" y más..." if len(dias_sin_carga) > 3 else ""))

    dev = calcular_desvio(df, operario, mes, anio)
    if not dev.empty:
        grandes = dev[abs(dev['Desvío (%)']) > 50]
        for _, row in grandes.iterrows():
            signo = "+" if row['Desvío (hs)'] > 0 else ""
            alertas.append(f"⚠️ {row['Tarea']}: desvío de {signo}{row['Desvío (hs)']} hs ({signo}{row['Desvío (%)']}%) vs. promedio histórico")

    df_disp = df_m[df_m['Tarea'].isin(TAREAS_DISPONIBILIDAD)][operario].sum()
    h_ina = df_m[df_m['Tarea'].isin(TAREAS_DESCUENTO_CAPACIDAD)][operario].sum()
    cap = capacidad_neta(anio, mes, feriados, h_ina)
    disp_pct = (df_disp / cap * 100) if cap > 0 else 0
    if disp_pct < 10 and cap > 0:
        alertas.append(f"🔴 Saturación crítica: solo {round(disp_pct, 1)}% de disponibilidad este mes")

    dias_sat, dias_tot = dias_saturados_mes(df, operario, anio, mes, feriados)
    if dias_tot > 0 and dias_sat / dias_tot > 0.5:
        alertas.append(f"🔴 Saturada {dias_sat} de {dias_tot} días hábiles del mes (más de la mitad).")

    return alertas

# ===== 10. MATRIZ DE COMPETENCIAS =====
def tiene_competencia(df_comp, integrante, tarea, subtarea):
    """Se asume que todos saben todo, salvo una excepción explícita 'No' en la hoja."""
    if df_comp is None or df_comp.empty:
        return True
    sub_norm = subtarea if subtarea and str(subtarea).strip() not in ('', '—') else '—'
    match = df_comp[
        (df_comp['Integrante'] == integrante) &
        (df_comp['Tarea'] == tarea) &
        (df_comp['Subtarea'].astype(str) == str(sub_norm))
    ]
    if match.empty:
        return True
    valor = str(match.iloc[0].get('Sabe', 'Sí')).strip().lower()
    return valor not in ('no', 'false', '0')

# ===== 11. VISTA DIARIA DEL EQUIPO Y SATURACIÓN =====
def horas_cargadas_dia(df, integrante, fecha, excluir_index=None):
    if df.empty or integrante not in df.columns:
        return 0.0
    df_dia = df[df['Fecha'].dt.date == fecha]
    if excluir_index is not None and excluir_index in df_dia.index:
        df_dia = df_dia.drop(excluir_index)
    return round(df_dia[integrante].sum(), 1)

def dias_saturados_mes(df, operario, anio, mes, feriados):
    """Cuenta cuántos días hábiles del mes esa persona llegó/superó las 6hs."""
    ini, fin = get_rango_mes(anio, mes)
    dias_lab = pd.bdate_range(start=ini, end=fin, freq='C', holidays=feriados)
    if df.empty:
        return 0, len(dias_lab)
    count_sat = 0
    for d in dias_lab:
        if horas_cargadas_dia(df, operario, d.date()) >= HORAS_DIA_LABORAL:
            count_sat += 1
    return count_sat, len(dias_lab)

def foto_dia(df, fecha):
    df_dia = df[df['Fecha'].dt.date == fecha]
    rows = []
    for op in OPERARIOS_FIJOS:
        df_op = df_dia[df_dia[op] > 0]
        total = round(df_op[op].sum(), 1)
        margen = round(HORAS_DIA_LABORAL - total, 1)
        tareas_dia = df_op.apply(lambda r: etiqueta_tarea(r['Tarea'], r.get('Subtarea', '')), axis=1).tolist()
        estado, css_class = semaforo_dia(margen)
        rows.append({
            'Integrante': op, 'Horas Cargadas': total, 'Margen (hs)': max(margen, 0),
            'Tareas del día': ", ".join(tareas_dia) if tareas_dia else "Sin carga registrada",
            'Estado': estado, '_css': css_class
        })
    return pd.DataFrame(rows)

def sugerir_ayuda(df, fecha, tarea_objetivo, subtarea_objetivo, df_comp, excluir=None):
    df_dia = df[df['Fecha'].dt.date == fecha]
    candidatos = []
    for op in OPERARIOS_FIJOS:
        if op == excluir:
            continue
        total = df_dia[op].sum() if not df_dia.empty else 0
        margen = round(HORAS_DIA_LABORAL - total, 1)
        sabe = tiene_competencia(df_comp, op, tarea_objetivo, subtarea_objetivo)
        if margen > 0 and sabe:
            candidatos.append({'Integrante': op, 'Margen ese día (hs)': margen})
    if not candidatos:
        return pd.DataFrame()
    return pd.DataFrame(candidatos).sort_values('Margen ese día (hs)', ascending=False).reset_index(drop=True)

# ===== 12. PROYECCIÓN DE CAPACIDAD (NUEVO) =====
def demanda_vs_capacidad(df, anio_ref, mes_ref, feriados, meses_atras=6):
    """Demanda real (horas trabajadas, sin contar Disponible/Planificaciones) vs. capacidad neta total del equipo, mes a mes."""
    rows = []
    for i in range(meses_atras - 1, -1, -1):
        m = mes_ref - i; a = anio_ref
        if m <= 0: m += 12; a -= 1
        df_m = df[(df['Fecha'].dt.month == m) & (df['Fecha'].dt.year == a)]
        demanda_total = 0.0
        capacidad_total = 0.0
        for op in OPERARIOS_FIJOS:
            h_ina = df_m[df_m['Tarea'].isin(TAREAS_DESCUENTO_CAPACIDAD)][op].sum()
            cap = capacidad_neta(a, m, feriados, h_ina)
            h_disp = df_m[df_m['Tarea'].isin(TAREAS_DISPONIBILIDAD)][op].sum()
            demanda_total += (df_m[op].sum() - h_disp)
            capacidad_total += cap
        rows.append({'Mes': MESES_ES[m], 'Orden': meses_atras - 1 - i,
                     'Demanda (hs)': round(demanda_total, 1), 'Capacidad (hs)': round(capacidad_total, 1)})
    return pd.DataFrame(rows)

def proyectar_tendencia(df_dc, meses_a_futuro=3):
    """Ajuste lineal simple sobre la demanda histórica para proyectar meses futuros."""
    if len(df_dc) < 3:
        return None
    x = df_dc['Orden'].values
    y = df_dc['Demanda (hs)'].values
    pendiente, ordenada = np.polyfit(x, y, 1)
    x_futuro = np.arange(x.max() + 1, x.max() + 1 + meses_a_futuro)
    y_futuro = pendiente * x_futuro + ordenada
    return pendiente, y_futuro

# ===== 13. PDF =====
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

# ===== 14. GRÁFICO DE COMPOSICIÓN — BARRAS HORIZONTALES =====
def grafico_composicion(df_res, columna_valor, titulo):
    df_plot = df_res[df_res[columna_valor] > 0].sort_values(columna_valor, ascending=True)
    total = df_plot[columna_valor].sum()
    df_plot['Porcentaje'] = (df_plot[columna_valor] / total * 100).round(1) if total > 0 else 0
    df_plot['Etiqueta'] = df_plot.apply(lambda r: f"{r[columna_valor]} hs ({r['Porcentaje']}%)", axis=1)
    fig = px.bar(
        df_plot, x=columna_valor, y='Tarea', orientation='h',
        text='Etiqueta', title=titulo, color=columna_valor,
        color_continuous_scale=['#00b4d8', '#0077b6']
    )
    fig.update_traces(textposition='outside')
    fig.update_layout(coloraxis_showscale=False, height=max(300, 35 * len(df_plot)),
                       margin=dict(l=0, r=60, t=40, b=0), xaxis_title="Horas", yaxis_title="")
    return fig

# ===== 15. HEATMAP CALENDARIO DEL MES (NUEVO) =====
def heatmap_calendario_mes(df, anio, mes, feriados):
    ini, fin = get_rango_mes(anio, mes)
    dias_lab = pd.bdate_range(start=ini, end=fin, freq='C', holidays=feriados)
    z, etiquetas_dias = [], []
    for d in dias_lab:
        fila = [horas_cargadas_dia(df, op, d.date()) for op in OPERARIOS_FIJOS]
        z.append(fila)
        etiquetas_dias.append(d.strftime('%d/%m'))
    if not z:
        return None
    fig = go.Figure(data=go.Heatmap(
        z=z, x=OPERARIOS_FIJOS, y=etiquetas_dias,
        colorscale=[[0, '#2d6a4f'], [0.5, '#e9c46a'], [1, '#e63946']],
        zmin=0, zmax=HORAS_DIA_LABORAL, text=z, texttemplate="%{text}", showscale=True,
        colorbar=dict(title="hs/día")
    ))
    fig.update_layout(title="Calendario de Saturación del Mes (hs cargadas por día)",
                       height=max(400, 24 * len(etiquetas_dias)), margin=dict(l=0, r=0, t=40, b=0))
    return fig

# ===== 16. CONEXIÓN GOOGLE SHEETS (con manejo de errores explícito) =====
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
        if 'Subtarea' not in df.columns:
            df['Subtarea'] = ''
        return df, ws, None
    except Exception as e:
        return pd.DataFrame(), None, str(e)

def guardar_df(nombre, df):
    """Reescribe la hoja completa. Usar solo para ediciones/borrados/reset/competencias
    (operaciones poco frecuentes), NO para altas nuevas — ver agregar_filas()."""
    try:
        ws = conectar().worksheet(nombre); ws.clear()
        df_c = df.copy()
        for c in df_c.columns:
            if 'fecha' in c.lower():
                df_c[c] = pd.to_datetime(df_c[c], errors='coerce').dt.strftime('%d/%m/%Y')
        ws.update([df_c.columns.values.tolist()] + df_c.fillna('').astype(str).values.tolist())
        st.cache_data.clear()
        return True, None
    except Exception as e:
        return False, str(e)

def agregar_filas(nombre, filas, columnas):
    """Agrega filas nuevas SIN reescribir toda la hoja (evita pisar datos de otra
    persona que haya guardado casi al mismo tiempo)."""
    try:
        ws = conectar().worksheet(nombre)
        filas_fmt = []
        for fila_dict in filas:
            f = fila_dict.copy()
            for c in columnas:
                if 'fecha' in c.lower() and c in f:
                    v = f[c]
                    f[c] = pd.to_datetime(v, errors='coerce').strftime('%d/%m/%Y') if pd.notna(v) else ''
            filas_fmt.append([str(f.get(c, '')) for c in columnas])
        ws.append_rows(filas_fmt, value_input_option='USER_ENTERED')
        st.cache_data.clear()
        return True, None
    except Exception as e:
        return False, str(e)

def mostrar_error_tecnico(mensaje):
    if mensaje:
        with st.expander("⚠️ Ver detalle técnico del error"):
            st.code(mensaje)

@st.cache_data
def cargar_feriados_multi(anios):
    feriados = []
    errores = []
    for a in anios:
        try:
            df_f = pd.read_csv(f"feriados_{a}.csv")
            df_f.columns = df_f.columns.str.strip().str.lower()
            feriados += pd.to_datetime(df_f['fecha'], errors='coerce').dt.date.dropna().tolist()
        except Exception:
            errores.append(a)
            continue
    return feriados

FERIADOS = cargar_feriados_multi(ANIOS_FERIADOS_A_CARGAR)

# ===== 17. ESTADO INICIAL =====
COLUMNAS_BASE = ['Fecha', 'Tarea', 'Subtarea'] + OPERARIOS_FIJOS + ['Nota']
COLUMNAS_COMPETENCIAS = ['Integrante', 'Tarea', 'Subtarea', 'Sabe']

if 'cargas' not in st.session_state:
    df, _, err = cargar_hoja("Cargas")
    if not df.empty:
        if 'Subtarea' not in df.columns:
            df['Subtarea'] = ''
        st.session_state.cargas = df
    else:
        st.session_state.cargas = pd.DataFrame(columns=COLUMNAS_BASE)
    st.session_state.error_carga_inicial = err

if 'competencias' not in st.session_state:
    df_comp, _, err_comp = cargar_hoja("Competencias")
    st.session_state.competencias = df_comp if not df_comp.empty else pd.DataFrame(columns=COLUMNAS_COMPETENCIAS)

if 'usuario_actual' not in st.session_state:
    st.session_state.usuario_actual = None
if 'editando_idx' not in st.session_state:
    st.session_state.editando_idx = None

# ===== 18. LOGIN CON CONTRASEÑA =====
if st.session_state.usuario_actual is None:
    st.title("🏛️ CRM Grupo Pressacco")
    u = st.selectbox("Usuario:", ["Seleccionar..."] + OPERARIOS_FIJOS + ["Admin - Ver todo"])
    pwd = st.text_input("Contraseña:", type="password")
    if st.button("Ingresar") and u != "Seleccionar...":
        claves = st.secrets.get("passwords", {})
        clave_correcta = claves.get(u)
        if clave_correcta is None:
            st.error("No hay contraseña configurada para este usuario. Pedile al Admin que la agregue en Secrets.")
        elif pwd == clave_correcta:
            st.session_state.usuario_actual = u; st.rerun()
        else:
            st.error("Contraseña incorrecta.")
    st.stop()

# ===== 19. ALERTA MENSUAL =====
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

# ===== 20. NAVEGACIÓN =====
es_admin = st.session_state.usuario_actual == "Admin - Ver todo"
menu_opciones = ["📊 Panel de Control", "➕ Cargar Horas", "📁 Carga Masiva", "🧩 Competencias",
                  "📚 Manual", "📜 Protocolo", "⚙️ Reset"] if es_admin \
    else ["📊 Panel de Control", "➕ Cargar Mis Horas", "📚 Manual", "📜 Protocolo"]
menu = st.sidebar.radio("Navegación", menu_opciones)
if st.sidebar.button("Cerrar Sesión"):
    st.session_state.clear(); st.rerun()

# ===== 21. PANEL DE CONTROL =====
if "Panel de Control" in menu:
    st.title("📊 Análisis de Capacidad y Eficiencia")

    c1, c2, c3 = st.columns([1, 1, 2])
    with c1: anio = st.selectbox("Año", [2025, 2026, 2027], index=1)
    with c2: mes = st.selectbox("Mes", list(range(1, 13)), format_func=lambda x: MESES_ES[x], index=hoy.month - 1)
    with c3:
        p_sel = st.selectbox("Integrante:", OPERARIOS_FIJOS) if es_admin else st.session_state.usuario_actual

    df_p = st.session_state.cargas.copy()
    df_p['Fecha'] = pd.to_datetime(df_p['Fecha'], errors='coerce')
    if 'Subtarea' not in df_p.columns:
        df_p['Subtarea'] = ''
    df_act = df_p[(df_p['Fecha'].dt.month == mes) & (df_p['Fecha'].dt.year == anio)]
    df_comp = st.session_state.competencias

    tabs_labels = ["🟢 Resumen", "🔍 Análisis Individual", "🗂️ Comparativas"] + (["🌐 Equipo", "📈 Proyección"] if es_admin else [])
    tabs = st.tabs(tabs_labels)

    # ---------- PESTAÑA 1: RESUMEN ----------
    with tabs[0]:
        h_ina = df_act[df_act['Tarea'].isin(TAREAS_DESCUENTO_CAPACIDAD)][p_sel].sum()
        cap = capacidad_neta(anio, mes, FERIADOS, h_ina)
        total_cargado = round(df_act[p_sel].sum(), 1)
        h_disp = df_act[df_act['Tarea'].isin(TAREAS_DISPONIBILIDAD)][p_sel].sum()
        disp_pct = round((h_disp / cap * 100) if cap > 0 else 0, 1)
        util_pct = round(((total_cargado - h_disp) / cap * 100) if cap > 0 else 0, 1)
        dias_sat, dias_tot = dias_saturados_mes(df_p, p_sel, anio, mes, FERIADOS)

        k1, k2, k3, k4, k5 = st.columns(5)
        with k1: st.markdown(f'<div class="kpi-card"><h2>{total_cargado}</h2><p>Horas Cargadas</p></div>', unsafe_allow_html=True)
        with k2: st.markdown(f'<div class="kpi-card"><h2>{cap}</h2><p>Capacidad Neta</p></div>', unsafe_allow_html=True)
        with k3: st.markdown(f'<div class="kpi-card"><h2>{util_pct}%</h2><p>Utilización Real</p></div>', unsafe_allow_html=True)
        with k4: st.markdown(f'<div class="kpi-card"><h2>{disp_pct}%</h2><p>Disponibilidad {semaforo_estado(disp_pct)}</p></div>', unsafe_allow_html=True)
        with k5: st.markdown(f'<div class="kpi-card"><h2>{dias_sat}/{dias_tot}</h2><p>Días Saturados del Mes</p></div>', unsafe_allow_html=True)

        alertas = generar_alertas(df_p, p_sel, anio, mes, FERIADOS)
        if alertas:
            st.subheader("🚨 Alertas Automáticas")
            for a in alertas:
                st.markdown(f'<div class="alerta-box">{a}</div>', unsafe_allow_html=True)

        st.divider()

        st.subheader("📅 Vista Diaria del Equipo")
        st.caption("Elegí un día puntual para ver quién estaba saturado y quién podía dar una mano.")
        fecha_vista = st.date_input("Día a analizar", value=hoy.date(), key="fecha_vista_diaria")
        df_foto = foto_dia(df_p, fecha_vista)

        for _, row in df_foto.iterrows():
            st.markdown(
                f'<div class="dia-card {row["_css"]}"><b>{row["Integrante"]}</b> — {row["Estado"]} '
                f'| Cargado: {row["Horas Cargadas"]} hs | Margen: {row["Margen (hs)"]} hs<br>'
                f'<small>{row["Tareas del día"]}</small></div>',
                unsafe_allow_html=True
            )

        saturados = df_foto[df_foto['Estado'] == "🔴 Saturado"]
        if not saturados.empty and es_admin:
            st.markdown("**¿Quién puede ayudar?**")
            for _, row_sat in saturados.iterrows():
                op_sat = row_sat['Integrante']
                df_dia_op = df_p[(df_p['Fecha'].dt.date == fecha_vista) & (df_p[op_sat] > 0)]
                if df_dia_op.empty:
                    continue
                tarea_sat = df_dia_op.iloc[0]['Tarea']
                subtarea_sat = df_dia_op.iloc[0].get('Subtarea', '')
                candidatos = sugerir_ayuda(df_p, fecha_vista, tarea_sat, subtarea_sat, df_comp, excluir=op_sat)
                etiqueta_sat = etiqueta_tarea(tarea_sat, subtarea_sat)
                with st.expander(f"🔴 {op_sat} está saturada en «{etiqueta_sat}» — ¿quién puede ayudar?"):
                    if candidatos.empty:
                        st.warning(f"Nadie disponible **con la competencia de «{etiqueta_sat}»** ese día. "
                                   f"Esto puede indicar una necesidad de capacitación cruzada o de contratación.")
                    else:
                        st.dataframe(candidatos, use_container_width=True, hide_index=True)

    # ---------- PESTAÑA 2: ANÁLISIS INDIVIDUAL ----------
    with tabs[1]:
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
                try:
                    styled = dev.style.map(colorear, subset=['Desvío (hs)', 'Desvío (%)'])
                except AttributeError:
                    styled = dev.style.applymap(colorear, subset=['Desvío (hs)', 'Desvío (%)'])
                st.dataframe(styled, use_container_width=True, hide_index=True)
            with col_g:
                fig_dev = px.bar(dev, x='Desvío (hs)', y='Tarea', orientation='h',
                                 color='Desvío (hs)', color_continuous_scale=['#e63946', '#adb5bd', '#2d6a4f'],
                                 title="Desvío en horas (actual vs. prom. histórico)")
                fig_dev.update_layout(coloraxis_showscale=False, height=350, margin=dict(l=0, r=0, t=40, b=0))
                st.plotly_chart(fig_dev, use_container_width=True)
        else:
            st.info("No hay datos históricos suficientes para calcular desvíos (se necesitan al menos 2 meses previos).")

        st.divider()
        st.subheader(f"📈 Tendencia Histórica — {p_sel} (últimos 6 meses)")
        df_tend = tendencia_historica(df_p, p_sel, mes, anio)
        if not df_tend.empty:
            orden_meses = df_tend.sort_values('Orden', ascending=False)['Mes'].unique().tolist()
            fig_tend = px.bar(df_tend, x='Mes', y='Horas', color='Tarea',
                              category_orders={'Mes': orden_meses},
                              title="Horas por tarea/subtarea — evolución mensual")
            fig_tend.update_layout(barmode='stack', height=380, margin=dict(l=0, r=0, t=40, b=0))
            st.plotly_chart(fig_tend, use_container_width=True)

        st.divider()
        st.subheader(f"📅 Distribución Semanal — {MESES_ES[mes]} {anio}")
        df_sem = distribucion_semanal(df_p, p_sel, anio, mes)
        if not df_sem.empty:
            fig_sem = px.bar(df_sem, x='Semana', y=p_sel, color='Tarea',
                             barmode='stack', title="Carga por semana")
            fig_sem.update_layout(height=320, margin=dict(l=0, r=0, t=40, b=0))
            st.plotly_chart(fig_sem, use_container_width=True)
        else:
            st.info("Sin datos para distribución semanal.")

        st.divider()
        df_act2 = df_act.copy()
        df_act2['Etiqueta'] = df_act2.apply(lambda r: etiqueta_tarea(r['Tarea'], r.get('Subtarea', '')), axis=1)
        res_ind = df_act2.groupby('Etiqueta')[p_sel].sum().round(1).reset_index().rename(columns={'Etiqueta': 'Tarea'})
        res_graf = res_ind[(res_ind[p_sel] > 0) & (~res_ind['Tarea'].isin(TAREAS_DESCUENTO_CAPACIDAD))]
        if not res_graf.empty:
            st.subheader(f"📊 Composición de Horas — {p_sel}")
            col_g, col_m = st.columns([2, 1])
            with col_g:
                fig = grafico_composicion(res_graf, p_sel, f"Distribución por Tarea/Subtarea — {p_sel}")
                st.plotly_chart(fig, use_container_width=True)
            with col_m:
                if st.button("📥 PDF Mensual Individual"):
                    dat = [["Tarea/Subtarea", "Horas"]] + [[r['Tarea'], r[p_sel]] for _, r in res_ind.iterrows()] + [["TOTAL", total_cargado]]
                    st.download_button("Guardar Mensual",
                        generar_pdf_base(f"Reporte {p_sel}", f"{MESES_ES[mes]} {anio}", [("Detalle", dat)],
                                         incluir_grafico=res_ind.set_index('Tarea')[p_sel].to_dict()),
                        f"Mensual_{p_sel}.pdf")

    # ---------- PESTAÑA 3: COMPARATIVAS ----------
    with tabs[2]:
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
            df_m2 = df_m.copy()
            df_m2['Etiqueta'] = df_m2.apply(lambda r: etiqueta_tarea(r['Tarea'], r.get('Subtarea', '')), axis=1)
            hist_pdf[MESES_ES[m_c]] = df_m2.groupby('Etiqueta')[p_sel].sum().to_dict()

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
                                 [("Detalle", [["Tarea/Subtarea"] + meses_n] + rows)]),
                f"Trimestral_{p_sel}.pdf")

    # ---------- PESTAÑA 4: EQUIPO (solo admin) ----------
    if es_admin:
        with tabs[3]:
            st.subheader("🌐 Balance del Equipo")
            df_bal = balance_equipo(df_p, anio, mes, FERIADOS)
            st.dataframe(df_bal, use_container_width=True, hide_index=True)

            st.divider()
            fig_cal = heatmap_calendario_mes(df_p, anio, mes, FERIADOS)
            if fig_cal:
                st.plotly_chart(fig_cal, use_container_width=True)

            df_heat = df_act[~df_act['Tarea'].isin(TAREAS_DESCUENTO_CAPACIDAD)].copy()
            if not df_heat.empty:
                df_heat['Etiqueta'] = df_heat.apply(lambda r: etiqueta_tarea(r['Tarea'], r.get('Subtarea', '')), axis=1)
                heat_data = df_heat.groupby('Etiqueta')[OPERARIOS_FIJOS].sum().round(1)
                fig_heat = go.Figure(data=go.Heatmap(
                    z=heat_data.values, x=heat_data.columns.tolist(), y=heat_data.index.tolist(),
                    colorscale='Blues', text=heat_data.values, texttemplate="%{text}",
                    showscale=True
                ))
                fig_heat.update_layout(title="Heatmap de Horas por Tarea/Subtarea e Integrante",
                                       height=400, margin=dict(l=0, r=0, t=40, b=0))
                st.plotly_chart(fig_heat, use_container_width=True)

            df_eq = df_act[~df_act['Tarea'].isin(TAREAS_DESCUENTO_CAPACIDAD)].copy()
            df_eq['Etiqueta'] = df_eq.apply(lambda r: etiqueta_tarea(r['Tarea'], r.get('Subtarea', '')), axis=1)
            df_eq2 = df_eq.melt(id_vars=['Fecha', 'Tarea', 'Etiqueta'], value_vars=OPERARIOS_FIJOS, var_name='Op', value_name='Hs')
            res_eq = df_eq2.groupby('Etiqueta')['Hs'].sum().reset_index().rename(columns={'Etiqueta': 'Tarea'})
            res_eq_graf = res_eq.rename(columns={'Hs': 'Total'})
            if not res_eq_graf.empty and res_eq_graf['Total'].sum() > 0:
                fig_g = grafico_composicion(res_eq_graf, 'Total', "Total Horas Equipo (Netas)")
                st.plotly_chart(fig_g, use_container_width=True)

            if st.button("📥 PDF Global"):
                hist_g = {}
                for i in range(3):
                    m_c = mes - i; a_c = anio
                    if m_c <= 0: m_c += 12; a_c -= 1
                    df_m_g = df_p[(df_p['Fecha'].dt.month == m_c) & (df_p['Fecha'].dt.year == a_c)].copy()
                    df_m_g['Etiqueta'] = df_m_g.apply(lambda r: etiqueta_tarea(r['Tarea'], r.get('Subtarea', '')), axis=1)
                    df_net = df_m_g[~df_m_g['Tarea'].isin(TAREAS_DESCUENTO_CAPACIDAD)]
                    hist_g[MESES_ES[m_c]] = df_net.groupby('Etiqueta')[OPERARIOS_FIJOS].sum().sum(axis=1).to_dict()
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
                                     [("Consolidado", [["Tarea/Subtarea"] + mg] + rg)],
                                     incluir_grafico=res_eq.set_index('Tarea')['Hs'].to_dict()),
                    "Global_Neto.pdf")

        # ---------- PESTAÑA 5: PROYECCIÓN DE CAPACIDAD (NUEVO, solo admin) ----------
        with tabs[4]:
            st.subheader("📈 Proyección de Capacidad — Demanda vs. Capacidad del Equipo")
            st.caption("Pensado para la reunión de contratación: ¿la demanda está creciendo más rápido que la capacidad?")
            meses_hist_sel = st.slider("Meses de historia a considerar", 3, 12, 6)
            df_dc = demanda_vs_capacidad(df_p, anio, mes, FERIADOS, meses_atras=meses_hist_sel)

            if df_dc.empty or df_dc['Demanda (hs)'].sum() == 0:
                st.info("No hay datos suficientes todavía para proyectar (necesitás al menos 3 meses cargados).")
            else:
                fig_proy = go.Figure()
                fig_proy.add_trace(go.Scatter(x=df_dc['Mes'], y=df_dc['Demanda (hs)'], name='Demanda real',
                                               mode='lines+markers', line=dict(color='#e63946', width=3)))
                fig_proy.add_trace(go.Scatter(x=df_dc['Mes'], y=df_dc['Capacidad (hs)'], name='Capacidad neta del equipo',
                                               mode='lines+markers', line=dict(color='#0077b6', width=3, dash='dot')))

                resultado = proyectar_tendencia(df_dc, meses_a_futuro=3)
                if resultado:
                    pendiente, y_futuro = resultado
                    ultima_cap = df_dc['Capacidad (hs)'].iloc[-1]
                    meses_futuros = [f"+{i+1}" for i in range(len(y_futuro))]
                    fig_proy.add_trace(go.Scatter(x=meses_futuros, y=y_futuro, name='Demanda proyectada',
                                                   mode='lines+markers', line=dict(color='#e63946', width=2, dash='dash')))

                    brecha = round(y_futuro[-1] - ultima_cap, 1)
                    tendencia_txt = "creciendo" if pendiente > 0 else ("estable" if abs(pendiente) < 1 else "bajando")
                    st.metric(f"Demanda proyectada en 3 meses vs. capacidad actual",
                              f"{round(y_futuro[-1],1)} hs", delta=f"{brecha:+.1f} hs vs. capacidad de hoy")
                    if brecha > 0:
                        st.warning(f"⚠️ Si la tendencia actual se mantiene (demanda {tendencia_txt} a razón de "
                                   f"{round(pendiente,1)} hs/mes), en 3 meses la demanda superaría la capacidad "
                                   f"neta actual por **{brecha} hs** — señal de que podría hacer falta sumar personal.")
                    else:
                        st.success(f"✅ La demanda está {tendencia_txt}. Con la capacidad actual del equipo alcanzaría "
                                   f"para absorber la proyección de los próximos 3 meses.")

                fig_proy.update_layout(title="Demanda real vs. Capacidad neta — histórico y proyección",
                                        height=420, margin=dict(l=0, r=0, t=40, b=0), yaxis_title="Horas")
                st.plotly_chart(fig_proy, use_container_width=True)
                st.dataframe(df_dc[['Mes', 'Demanda (hs)', 'Capacidad (hs)']], use_container_width=True, hide_index=True)

# ===== 22. CARGA MASIVA =====
elif "Carga Masiva" in menu:
    st.title("📁 Distribución Masiva")

    t_m = st.selectbox("Tarea", list(COLORES_TAREAS.keys()), key="tarea_masiva")
    subs_m = SUBTAREAS.get(t_m, [])
    st_m = st.selectbox("Subtarea", subs_m, key="subtarea_masiva") if subs_m else None

    with st.form("fm"):
        u_m = st.selectbox("Persona", OPERARIOS_FIJOS)
        st.write(f"**Tarea:** {t_m}" + (f" — {st_m}" if st_m else ""))
        f_i = st.date_input("Desde"); f_f = st.date_input("Hasta")
        h_t = st.number_input("Horas Totales", min_value=0.0, max_value=float(HORAS_MAX_DIA) * 60)
        if st.form_submit_button("Ejecutar"):
            rgo = pd.bdate_range(start=f_i, end=f_f, freq='C', holidays=FERIADOS)
            if len(rgo) > 0:
                h_d = round(h_t / len(rgo), 2)
                df_check = st.session_state.cargas.copy()
                df_check['Fecha'] = pd.to_datetime(df_check['Fecha'], errors='coerce')
                dias_saturados = []
                filas = []
                for d in rgo:
                    ya_cargado = horas_cargadas_dia(df_check, u_m, d.date())
                    if ya_cargado + h_d > HORAS_DIA_LABORAL:
                        dias_saturados.append(d.strftime('%d/%m'))
                    fila = {'Fecha': d, 'Tarea': t_m, 'Subtarea': st_m if st_m else '—', 'Nota': 'Masiva'}
                    for o in OPERARIOS_FIJOS: fila[o] = h_d if o == u_m else 0
                    filas.append(fila)
                if dias_saturados:
                    st.warning(f"⚠️ {u_m} va a superar las {HORAS_DIA_LABORAL} hs diarias en: {', '.join(dias_saturados[:5])}"
                               + (" y más..." if len(dias_saturados) > 5 else "") + ". Se guardó igual, revisá si corresponde.")
                st.session_state.cargas = pd.concat([st.session_state.cargas, pd.DataFrame(filas)], ignore_index=True)
                ok, err = agregar_filas("Cargas", filas, COLUMNAS_BASE)
                if ok:
                    st.success("✅ Guardado."); time.sleep(1.5); st.rerun()
                else:
                    st.error("No se pudo guardar en Google Sheets. Los datos quedaron en tu sesión pero no se sincronizaron.")
                    mostrar_error_tecnico(err)

# ===== 23. CARGA INDIVIDUAL (con edición) =====
elif "Cargar" in menu:
    st.title("➕ Registro de Horas")
    u_c = st.session_state.usuario_actual if not es_admin else st.selectbox("Persona:", OPERARIOS_FIJOS)

    f_t = st.selectbox("Tarea", list(COLORES_TAREAS.keys()), key="tarea_sel")
    subs_disponibles = SUBTAREAS.get(f_t, [])
    if subs_disponibles:
        f_sub = st.selectbox("Subtarea", subs_disponibles, key="subtarea_sel")
    else:
        f_sub = "—"
        st.caption("Esta tarea no tiene subtareas.")

    with st.form("fi"):
        f_fecha = st.date_input("Fecha", value=datetime.now())
        st.write(f"**Tarea seleccionada:** {f_t}" + (f" — {f_sub}" if f_sub != "—" else ""))
        f_h = st.number_input("Horas", step=0.5, value=6.0, min_value=0.5, max_value=float(HORAS_MAX_DIA))
        f_n = st.text_input("Nota")

        if st.form_submit_button("Guardar"):
            df_check = st.session_state.cargas.copy()
            df_check['Fecha'] = pd.to_datetime(df_check['Fecha'], errors='coerce')
            if 'Subtarea' not in df_check.columns:
                df_check['Subtarea'] = ''

            duplicado = df_check[
                (df_check['Fecha'].dt.date == f_fecha) &
                (df_check['Tarea'] == f_t) &
                (df_check['Subtarea'].astype(str) == str(f_sub)) &
                (df_check[u_c] > 0)
            ]
            if not duplicado.empty:
                st.warning(f"⚠️ Ya existe una carga para {u_c} el {f_fecha.strftime('%d/%m/%Y')} en '{f_t} — {f_sub}'. Usá 'Editar' en el historial si querés cambiarla.")
            else:
                ya_cargado_hoy = horas_cargadas_dia(df_check, u_c, f_fecha)
                total_dia = round(ya_cargado_hoy + f_h, 1)
                if total_dia > HORAS_DIA_LABORAL:
                    st.warning(f"⚠️ {u_c} va a quedar con {total_dia} hs cargadas el {f_fecha.strftime('%d/%m/%Y')} "
                               f"(máximo habitual: {HORAS_DIA_LABORAL} hs). Se guarda igual, revisá si corresponde.")

                nueva = {'Fecha': f_fecha, 'Tarea': f_t, 'Subtarea': f_sub, 'Nota': f_n}
                for op in OPERARIOS_FIJOS: nueva[op] = f_h if op == u_c else 0
                st.session_state.cargas = pd.concat([st.session_state.cargas, pd.DataFrame([nueva])], ignore_index=True)
                ok, err = agregar_filas("Cargas", [nueva], COLUMNAS_BASE)
                if ok:
                    st.success(f"✅ Guardado: {f_t}{' — ' + f_sub if f_sub != '—' else ''} | {f_h} hs")
                    time.sleep(1); st.rerun()
                else:
                    st.error("No se pudo sincronizar con Google Sheets. Quedó guardado solo en tu sesión actual.")
                    mostrar_error_tecnico(err)

    st.divider()
    mes_f = st.selectbox("Historial Mes:", list(range(1, 13)), format_func=lambda x: MESES_ES[x], index=hoy.month - 1)
    df_h = st.session_state.cargas.copy()
    df_h['Fecha'] = pd.to_datetime(df_h['Fecha'], errors='coerce')
    if 'Subtarea' not in df_h.columns:
        df_h['Subtarea'] = ''
    df_f = df_h[(df_h[u_c] > 0) & (df_h['Fecha'].dt.month == mes_f)]
    if not df_f.empty:
        st.info(f"**Total {MESES_ES[mes_f]}:** {round(df_f[u_c].sum(), 1)} hs")
        for i, r in df_f.sort_values('Fecha', ascending=False).iterrows():
            sub_label = f" — {r['Subtarea']}" if str(r.get('Subtarea', '')).strip() and r.get('Subtarea', '') != '—' else ''
            c1, c2, c3 = st.columns([5, 1, 1])
            c1.write(f"📅 {r['Fecha'].strftime('%d/%m/%Y')} | {r['Tarea']}{sub_label} | {r[u_c]} hs | {r['Nota']}")
            if c2.button("Editar", key=f"edit_{i}"):
                st.session_state.editando_idx = i; st.rerun()
            if c3.button("Eliminar", key=f"del_{i}"):
                st.session_state.cargas.drop(i, inplace=True)
                ok, err = guardar_df("Cargas", st.session_state.cargas)
                if ok:
                    st.rerun()
                else:
                    st.error("No se pudo eliminar en Google Sheets.")
                    mostrar_error_tecnico(err)

    # ----- FORMULARIO DE EDICIÓN -----
    if st.session_state.editando_idx is not None and st.session_state.editando_idx in st.session_state.cargas.index:
        idx = st.session_state.editando_idx
        row = st.session_state.cargas.loc[idx]
        st.divider()
        st.markdown("### ✏️ Editando carga")
        with st.form("form_editar"):
            e_fecha = st.date_input("Fecha", value=pd.to_datetime(row['Fecha']).date())
            lista_tareas = list(COLORES_TAREAS.keys())
            e_tarea = st.selectbox("Tarea", lista_tareas, index=lista_tareas.index(row['Tarea']) if row['Tarea'] in lista_tareas else 0)
            subs_e = SUBTAREAS.get(e_tarea, [])
            if subs_e:
                e_sub = st.selectbox("Subtarea", subs_e, index=subs_e.index(row['Subtarea']) if row['Subtarea'] in subs_e else 0)
            else:
                e_sub = "—"
            e_horas = st.number_input("Horas", step=0.5, value=float(row[u_c]) if row[u_c] else 0.5, min_value=0.5, max_value=float(HORAS_MAX_DIA))
            e_nota = st.text_input("Nota", value=str(row.get('Nota', '')))
            cg1, cg2 = st.columns(2)
            guardar_edit = cg1.form_submit_button("💾 Guardar cambios")
            cancelar_edit = cg2.form_submit_button("Cancelar")

            if guardar_edit:
                st.session_state.cargas.loc[idx, 'Fecha'] = e_fecha
                st.session_state.cargas.loc[idx, 'Tarea'] = e_tarea
                st.session_state.cargas.loc[idx, 'Subtarea'] = e_sub
                st.session_state.cargas.loc[idx, u_c] = e_horas
                st.session_state.cargas.loc[idx, 'Nota'] = e_nota
                ok, err = guardar_df("Cargas", st.session_state.cargas)
                if ok:
                    st.success("✅ Carga actualizada.")
                    st.session_state.editando_idx = None
                    time.sleep(1); st.rerun()
                else:
                    st.error("No se pudo guardar el cambio en Google Sheets.")
                    mostrar_error_tecnico(err)
            if cancelar_edit:
                st.session_state.editando_idx = None; st.rerun()

# ===== 24. COMPETENCIAS (solo admin) =====
elif "Competencias" in menu:
    st.title("🧩 Matriz de Competencias")
    st.caption("Por defecto se asume que TODOS saben hacer TODAS las tareas. Acá marcás únicamente las "
               "EXCEPCIONES: quién NO sabe hacer determinada tarea/subtarea.")

    df_comp = st.session_state.competencias.copy()

    todas_combos = []
    for tarea, subs in SUBTAREAS.items():
        if subs:
            for s in subs:
                todas_combos.append((tarea, s))
        else:
            todas_combos.append((tarea, "—"))

    op_edit = st.selectbox("Integrante", OPERARIOS_FIJOS)
    st.write(f"Desmarcá lo que **{op_edit}** NO sabe hacer:")

    cambios = {}
    cols = st.columns(2)
    for i, (tarea, sub) in enumerate(todas_combos):
        etiqueta = etiqueta_tarea(tarea, sub)
        ya_no_sabe = not tiene_competencia(df_comp, op_edit, tarea, sub)
        with cols[i % 2]:
            sabe_check = st.checkbox(etiqueta, value=not ya_no_sabe, key=f"comp_{op_edit}_{tarea}_{sub}")
        cambios[(tarea, sub)] = sabe_check

    if st.button("💾 Guardar Competencias"):
        filas_no = []
        for (tarea, sub), sabe in cambios.items():
            if not sabe:
                filas_no.append({'Integrante': op_edit, 'Tarea': tarea, 'Subtarea': sub, 'Sabe': 'No'})

        df_otros = df_comp[df_comp['Integrante'] != op_edit] if not df_comp.empty else pd.DataFrame(columns=COLUMNAS_COMPETENCIAS)
        df_nueva = pd.concat([df_otros, pd.DataFrame(filas_no, columns=COLUMNAS_COMPETENCIAS)], ignore_index=True)
        st.session_state.competencias = df_nueva
        ok, err = guardar_df("Competencias", df_nueva)
        if ok:
            st.success(f"✅ Competencias de {op_edit} actualizadas.")
        else:
            st.error("No se pudo guardar. Verificá que exista la hoja 'Competencias' en el Google Sheet "
                     "con columnas: Integrante, Tarea, Subtarea, Sabe.")
            mostrar_error_tecnico(err)
        time.sleep(1.2); st.rerun()

# ===== 25. MANUAL =====
elif "Manual" in menu:
    st.title("📚 Manual de Uso — CRM Capacidad Instalada")
    st.markdown("**Versión 5.0** | Con autenticación, guardado seguro, proyección de capacidad y edición de cargas.")
    st.divider()

    with st.expander("🎯 1. ¿Qué es el CRM?", expanded=False):
        st.markdown("""
El CRM registra y analiza cómo se distribuye el tiempo del equipo: subtareas, Vista Diaria, matriz de
competencias y ahora también una **Proyección de Capacidad** para saber si conviene contratar.
        """)

    with st.expander("🔐 2. Acceso y contraseñas", expanded=False):
        st.markdown("""
Cada usuario (incluido Admin) necesita una contraseña, configurada por quien administra la app en
`Secrets` (`st.secrets`), bajo la sección `[passwords]`. Sin eso configurado, nadie puede entrar.
        """)

    with st.expander("🧩 3. Matriz de Competencias", expanded=False):
        st.markdown("""
Por defecto se asume que todos saben hacer todas las tareas. Solo hay que marcar las excepciones
(ej: "Maximiliano no sabe Sueldos — Liquidación") en **🧩 Competencias** (solo Admin). Esto se usa
automáticamente en la Vista Diaria para sugerir reemplazos reales.
        """)

    with st.expander("📅 4. Vista Diaria del Equipo", expanded=False):
        st.markdown("""
En **📊 Panel de Control → 🟢 Resumen**, elegí una fecha puntual y vas a ver el estado de las 4 personas
ese día, y si alguien está saturado, quién puede ayudar (según disponibilidad y competencia).
        """)

    with st.expander("📈 5. Proyección de Capacidad", expanded=False):
        st.markdown("""
En **📊 Panel de Control → 📈 Proyección** (solo Admin), se compara la demanda real de horas del equipo
contra la capacidad neta total, mes a mes, y se proyecta la tendencia a 3 meses. Sirve como base numérica
para decidir si conviene sumar personal.
        """)

    with st.expander("✍️ 6. Cargar, editar y eliminar horas", expanded=False):
        st.markdown("""
1. Ir a **➕ Cargar Mis Horas**, completar Fecha/Tarea/Subtarea/Horas/Nota y **Guardar**.
2. En el historial de abajo, cada carga tiene botones **Editar** y **Eliminar**.
3. Si con una carga superás las 6 hs del día (sumando todo lo ya cargado), el sistema avisa pero
   permite guardar igual, por si corresponde a una excepción real.
        """)

    with st.expander("🗓️ 7. Feriados", expanded=False):
        st.markdown(f"""
La app carga feriados de varios archivos `feriados_AAAA.csv` (actualmente: {', '.join(str(a) for a in ANIOS_FERIADOS_A_CARGAR)}).
Cuando empiece un año nuevo que no esté en la lista, hay que subir el `feriados_AAAA.csv` correspondiente
y agregar el año a `ANIOS_FERIADOS_A_CARGAR` en el código.
        """)

# ===== 26. PROTOCOLO =====
elif "Protocolo" in menu:
    st.title("📜 Protocolo de Uso")
    if st.button("📥 Descargar Guía"):
        st.download_button("Guardar",
            generar_pdf_base("PROTOCOLO CRM", "Manual de Procedimientos", [], es_protocolo=True),
            "Protocolo_Pressacco.pdf")

# ===== 27. RESET (con backup obligatorio previo) =====
elif "Reset" in menu:
    st.title("⚙️ Resetear Base")
    st.warning("⚠️ Esto borra TODOS los datos cargados. Descargá el backup antes de continuar.")

    csv_backup = st.session_state.cargas.to_csv(index=False).encode('utf-8')
    st.download_button("📥 Descargar backup CSV antes de borrar", csv_backup,
                        f"backup_cargas_{hoy.strftime('%Y%m%d_%H%M')}.csv", "text/csv")

    st.divider()
    if st.text_input("Escriba BORRAR") == "BORRAR":
        if st.button("Limpiar Todo"):
            ok, err = guardar_df("Cargas", pd.DataFrame(columns=COLUMNAS_BASE))
            if ok:
                st.session_state.cargas = pd.DataFrame(columns=COLUMNAS_BASE)
                st.success("✅ Base reseteada.")
                st.rerun()
            else:
                st.error("No se pudo resetear la base en Google Sheets.")
                mostrar_error_tecnico(err)
