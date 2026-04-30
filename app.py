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
    estilo_titulo = s['Title']; estilo_titulo.textColor = color_celeste
    estilo_cuerpo = s['Normal']; estilo_cuerpo.fontSize = 9; estilo_cuerpo.leading = 12
    estilo_subtitulo = s['Heading3']; estilo_subtitulo.textColor = color_celeste
    estilo_negrita = s['Normal']; estilo_negrita.fontSize = 9; estilo_negrita.fontName = 'Helvetica-Bold'

    story = [Paragraph("GRUPO PRESSACCO", estilo_titulo), Paragraph(titulo_doc, s['Heading2']), Paragraph(subtitulo, s['Normal']), Spacer(1, 15)]

    if es_protocolo:
        # Incorporación de todo el Manual de Procedimientos redactado
        secciones = [
            ("1. INTRODUCCIÓN Y FINALIDAD", "El objetivo principal de este CRM es transformar nuestra carga de trabajo en datos accionables. En un estudio contable, el tiempo es nuestro recurso más valioso; registrarlo nos permite medir la rentabilidad de cada proceso, detectar cuándo un integrante está saturado y planificar el crecimiento del equipo con bases sólidas. Elegimos sumar precisión para restar incertidumbre."),
            ("2. OBJETIVOS ESTRATÉGICOS", "• <b>Visibilidad:</b> Saber en qué tareas invertimos más tiempo.<br/>• <b>Equilibrio:</b> Evitar cuellos de botella y redistribuir tareas.<br/>• <b>Transparencia:</b> Tener un registro histórico claro frente a auditorías."),
            ("3. PASO A PASO: CARGA Y CONTROL", "<b>• Carga Diaria:</b> Registrar 6 horas diarias antes de las 15:00 hs.<br/><b>• Sinceridad:</b> Usar tareas reales o DISPONIBLE según corresponda.<br/><b>• Inasistencias:</b> Se cargan para restar capacidad neta automáticamente.<br/><b>• Autocontrol:</b> Verificar en el Historial que el total sume 6 hs.<br/><b>• Reunión de Equipo:</b> Analizar el PDF Trimestral grupalmente."),
            ("4. SEMÁFORO DE CAPACIDAD", "• 🟢 <b>Verde (>20%):</b> Espacio para nuevos proyectos.<br/>• 🟡 <b>Amarillo (10-20%):</b> Carga próxima al límite.<br/>• 🔴 <b>Rojo (<10%):</b> Saturación operativa.")
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

# (Las secciones de Conexión, Alertas y Login se mantienen idénticas para no borrar nada)

# ===== 7. PANEL DE CONTROL =====
if "Panel de Control" in menu:
    st.title("📊 Análisis y Autoevaluación")
    c1, c2, c3 = st.columns([1,1,2])
    with c1: anio = st.selectbox("Año", [2025, 2026, 2027], index=1)
    with c2: mes = st.selectbox("Mes", list(range(1,13)), format_func=lambda x: MESES_ES[x], index=datetime.now().month-1)
    with c3: p_sel = st.selectbox("Integrante Individual:", OPERARIOS_FIJOS) if st.session_state.usuario_actual == "Admin - Ver todo" else st.session_state.usuario_actual

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
        pdf_t = generar_pdf_base(f"Trimestral Detallado: {p_sel}", "Comparativa de horas registradas", [("Desvío por Tarea", [["Tarea"] + meses_n] + rows)])
        st.download_button("Guardar Trimestral", pdf_t, f"Trimestral_{p_sel}.pdf")

    st.divider()
    df_act = df_p[(df_p['Fecha'].dt.month == mes) & (df_p['Fecha'].dt.year == anio)]
    res_ind = df_act.groupby('Tarea')[p_sel].sum().round(1).reset_index()
    res_ind = res_ind[res_ind[p_sel]>0]
    
    if not res_ind.empty:
        col_p, col_m = st.columns([2,1])
        res_neta_grafico = res_ind[~res_ind['Tarea'].isin(TAREAS_DESCUENTO_CAPACIDAD)]
        with col_p: st.plotly_chart(px.pie(res_neta_grafico, values=p_sel, names='Tarea', color='Tarea', color_discrete_map=COLORES_TAREAS, title=f"Eficiencia Real (Sin Inasistencias) - {p_sel}"), use_container_width=True)
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
                pdf_m = generar_pdf_base(f"Reporte Mensual: {p_sel}", f"{MESES_ES[mes]} {anio}", [("Detalle de Jornada", datos_t)], incluir_grafico=res_ind.set_index('Tarea')[p_sel].to_dict())
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
            pdf_g = generar_pdf_base("REPORTE GLOBAL NETO", "Estudio Completo - Sin Inasistencias", [("Consolidado Productivo", [["Tarea"] + meses_g] + rows_g)], incluir_grafico=res_eq_neta.set_index('Tarea')['Hs'].to_dict())
            st.download_button("Guardar Reporte Global", pdf_g, "Global_Neto.pdf")

# (Sección Cargar Horas con filtros e historial intacta)

elif "Protocolo" in menu:
    st.title("📜 Protocolo de Uso - Grupo Pressacco")
    st.markdown("""
    ### 1. Finalidad y Objetivo
    Transformar nuestra carga de trabajo en **datos accionables**. El objetivo es medir capacidad, eficiencia y asegurar un equilibrio saludable en el equipo.
    
    ### 2. Paso a Paso: Carga y Control
    *   **Registro Diario:** Cada integrante debe registrar **6 horas diarias** antes de las **15:00 hs**.
    *   **Gestión de Inasistencias:** Cargarlas como tal; el sistema las descuenta automáticamente de la capacidad neta.
    *   **Autocontrol:** Verificar en el Historial que el total diario sume 6 hs.
    """)
    if st.button("📥 Descargar Guía Maestra (PDF)"):
        # PDF con el manual completo redactado
        pdf = generar_pdf_base("PROTOCOLO DE USO - CRM", "Manual de Procedimientos y Guía Completa", [], es_protocolo=True)
        st.download_button("Guardar Protocolo Maestro", pdf, "Protocolo_Pressacco.pdf")

# (Sección Carga Masiva y Reset intactas)
