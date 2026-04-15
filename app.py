import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import plotly.express as px
import plotly.graph_objects as go
import gspread
from google.oauth2.service_account import Credentials
from io import BytesIO
import base64
import time

st.set_page_config(page_title="CRM Capacidad Instalada", layout="wide")

# ===== CONFIGURACIÓN =====
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

MESES_ES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
    7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
}

DIAS_SEMANA_ES = {
    0: "Lunes", 1: "Martes", 2: "Miércoles", 3: "Jueves",
    4: "Viernes", 5: "Sábado", 6: "Domingo"
}

# ===== FUNCIÓN PARA PARSEAR FECHAS =====
def parsear_fecha_flexible(valor):
    if pd.isna(valor) or str(valor).strip() == '':
        return None
    try:
        return pd.to_datetime(str(valor).strip(), dayfirst=True, errors='raise').date()
    except:
        return None

# ===== CONEXIÓN GOOGLE SHEETS =====
@st.cache_resource
def conectar_sheets():
    scopes = ["https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scopes)
    client = gspread.authorize(creds)
    return client.open("CRM_Estudio_Datos")

@st.cache_data(ttl=60)
def cargar_hoja(nombre_hoja):
    try:
        ws = conectar_sheets().worksheet(nombre_hoja)
        df = pd.DataFrame(ws.get_all_records())
        df.columns = df.columns.str.strip()
        for col in df.columns:
            if 'fecha' in col.lower():
                df[col] = df[col].apply(parsear_fecha_flexible)
        return df, ws
    except:
        return pd.DataFrame(), None

def guardar_df(nombre_hoja, df):
    try:
        ws = conectar_sheets().worksheet(nombre_hoja)
        ws.clear()
        df_copy = df.copy()
        for col in df_copy.columns:
            if 'fecha' in col.lower():
                df_copy[col] = pd.to_datetime(df_copy[col], errors='coerce').dt.strftime('%d/%m/%Y')
        df_copy = df_copy.fillna('').astype(str)
        if df_copy.empty:
            ws.update([df_copy.columns.values.tolist()])
        else:
            ws.update([df_copy.columns.values.tolist()] + df_copy.values.tolist())
        st.cache_data.clear()
        return True
    except Exception as e:
        st.error(f"Error guardando en {nombre_hoja}: {e}")
        return False

# ===== FUNCIÓN GENERAR PDF =====
def generar_pdf_reporte(tipo, nombre, mes, anio, capacidad, total_horas, porcentaje, estado, df_tareas, df_equipo=None):
    from reportlab.lib.pagesizes import letter
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib import colors
    from io import BytesIO

    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter)
    styles = getSampleStyleSheet()
    story = []

    # Título
    titulo = f"Reporte {tipo.title()} - {nombre}" if tipo == "individual" else f"Reporte Equipo - {MESES_ES[mes]} {anio}"
    story.append(Paragraph(titulo, styles['Title']))
    story.append(Spacer(1, 12))

    # Resumen
    if tipo == "individual":
        data_resumen = [
            ["Métrica", "Valor"],
            ["Período", f"{MESES_ES[mes]} {anio}"],
            ["Total Cargado", f"{total_horas:.1f} hs"],
            ["% Tiempo Disponible", f"{porcentaje:.1f}%"],
            ["Estado", estado],
            ["Capacidad Real", f"{capacidad:.0f} hs"]
        ]
    else:
        data_resumen = [
            ["Métrica", "Valor"],
            ["Período", f"{MESES_ES[mes]} {anio}"],
            ["Total Equipo", f"{total_horas:.1f} hs"],
            ["% Disponible Promedio", f"{porcentaje:.1f}%"]
        ]

    tabla_resumen = Table(data_resumen, colWidths=[200, 200])
    tabla_resumen.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))
    story.append(tabla_resumen)
    story.append(Spacer(1, 20))

    # Tabla de tareas
    story.append(Paragraph("Distribución por Tarea", styles['Heading2']))
    data_tareas = [["Tarea", "Horas"]]
    for _, row in df_tareas.iterrows():
        data_tareas.append([row['Tarea'], f"{row['Horas']:.1f}"])

    tabla_tareas = Table(data_tareas, colWidths=[300, 100])
    tabla_tareas.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))
    story.append(tabla_tareas)

    # Si es reporte de equipo, agregar tabla del equipo con las columnas nuevas
    if tipo == "equipo" and df_equipo is not None and not df_equipo.empty:
        story.append(Spacer(1, 20))
        story.append(Paragraph("Resumen del Equipo", styles['Heading2']))

        # Usa las columnas que tenga el dataframe, sin hardcodear nombres
        cols_equipo = list(df_equipo.columns)
        data_equipo = [cols_equipo]
        for _, row in df_equipo.iterrows():
            data_equipo.append([str(row[col]) for col in cols_equipo])

        tabla_equipo = Table(data_equipo)
        tabla_equipo.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        story.append(tabla_equipo)

    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()

# ===== FUNCIONES =====
def calcular_dias_habiles(fecha_inicio, fecha_fin):
    dias = pd.bdate_range(start=fecha_inicio, end=fecha_fin, freq='C', holidays=FERIADOS)
    return len(dias)

def calcular_capacidad_mensual(anio, mes):
    inicio = datetime(anio, mes, 1)
    if mes == 12:
        fin = datetime(anio + 1, 1, 1) - timedelta(days=1)
    else:
        fin = datetime(anio, mes + 1, 1) - timedelta(days=1)
    dias_habiles = calcular_dias_habiles(inicio.date(), fin.date())
    return dias_habiles, dias_habiles * HORAS_DIA_LABORAL

def calcular_excepciones_mes(operario, anio, mes, df_excepciones):
    if df_excepciones.empty:
        return 0
    df_exc = df_excepciones.copy()
    df_exc['Fecha Inicio'] = pd.to_datetime(df_exc['Fecha Inicio'], errors='coerce')
    df_exc['Fecha Fin'] = pd.to_datetime(df_exc['Fecha Fin'], errors='coerce')
    df_exc = df_exc[df_exc['Operario'] == operario]
    total_horas_exc = 0
    inicio_mes = datetime(anio, mes, 1).date()
    if mes == 12:
        fin_mes = datetime(anio + 1, 1, 1).date() - timedelta(days=1)
    else:
        fin_mes = datetime(anio, mes + 1, 1).date() - timedelta(days=1)
    for _, row in df_exc.iterrows():
        if pd.notna(row['Fecha Inicio']) and pd.notna(row['Fecha Fin']):
            inicio_exc = row['Fecha Inicio'].date()
            fin_exc = row['Fecha Fin'].date()
            inicio_efectivo = max(inicio_exc, inicio_mes)
            fin_efectivo = min(fin_exc, fin_mes)
            if inicio_efectivo <= fin_efectivo:
                dias_exc = calcular_dias_habiles(inicio_efectivo, fin_efectivo)
                total_horas_exc += dias_exc * HORAS_DIA_LABORAL
    return total_horas_exc

# ===== SESSION STATE =====
if 'cargas' not in st.session_state:
    df, _ = cargar_hoja("Cargas")
    if df.empty:
        df = pd.DataFrame(columns=['Fecha', 'Tarea'] + OPERARIOS_FIJOS + ['Nota'])
    st.session_state.cargas = df

if 'excepciones' not in st.session_state:
    df, _ = cargar_hoja("Excepciones")
    if df.empty:
        df = pd.DataFrame(columns=['Operario', 'Fecha Inicio', 'Fecha Fin', 'Motivo', 'Horas'])
    st.session_state.excepciones = df

# ===== LOGIN =====
if 'usuario_actual' not in st.session_state:
    st.session_state.usuario_actual = None

if st.session_state.usuario_actual is None:
    st.title("CRM Capacidad Instalada - Estudio")
    st.subheader("Seleccioná tu usuario para ingresar")
    usuario = st.selectbox("Soy:", ["Seleccionar..."] + OPERARIOS_FIJOS + ["Admin - Ver todo"])
    if st.button("Ingresar") and usuario!= "Seleccionar...":
        st.session_state.usuario_actual = usuario
        st.rerun()
    st.stop()

# ===== SIDEBAR =====
col_user, col_logout = st.sidebar.columns([3,1])
with col_user:
    st.sidebar.success(f"Usuario: **{st.session_state.usuario_actual}**")
with col_logout:
    if st.sidebar.button("Salir"):
        st.session_state.clear()
        st.rerun()

if st.session_state.usuario_actual == "Admin - Ver todo":
    menu = st.sidebar.radio("Menú", ["Panel de Control", "Cargar Horas", "Carga Masiva", "Excepciones", "Exportar Excel", "Resetear Datos"])
else:
    menu = st.sidebar.radio("Menú", ["Panel de Control", "Cargar Mis Horas"])

# ===== PANEL DE CONTROL =====
if menu == "Panel de Control":
    st.title("Panel de Control - Ocupación")

    # Definimos qué tareas suman al "Tiempo Disponible/No Productivo"
    TAREAS_DISPONIBLE = [
        "DISPONIBLE",
        "PLANIFICACIONES/ORGANIZACIÓN/PROCEDIMIENTO S/INFORMES",
        "INASISTENCIA POR EXAMEN O TRAMITE"
    ]

    col1, col2 = st.columns(2)
    with col1:
        anio = st.selectbox("Año", [2025, 2026, 2027], index=1)
    with col2:
        mes = st.selectbox("Mes", list(range(1,13)),
                           format_func=lambda x: MESES_ES[x],
                           index=3)

    dias_habiles, capacidad_base_mes = calcular_capacidad_mensual(anio, mes)
    st.write(f"**{MESES_ES[mes]} {anio} - {dias_habiles} días hábiles | Capacidad base del mes: {capacidad_base_mes}hs**")

    df_mes = st.session_state.cargas.copy()
    if not df_mes.empty:
        df_mes['Fecha'] = pd.to_datetime(df_mes['Fecha'], errors='coerce')
        mask = (df_mes['Fecha'].dt.month == mes) & (df_mes['Fecha'].dt.year == anio)
        df_mes = df_mes[mask]

    if df_mes.empty:
        st.info("Cargá horas para ver el gráfico")
    else:
        df_melt = df_mes.melt(id_vars=['Fecha', 'Tarea', 'Nota'],
                              value_vars=OPERARIOS_FIJOS,
                              var_name='Operario', value_name='Horas')
        df_melt = df_melt[df_melt['Horas'] > 0]

        st.subheader("Ocupación Individual")
        if st.session_state.usuario_actual == "Admin - Ver todo":
            persona_seleccionada = st.selectbox("Seleccionar persona", OPERARIOS_FIJOS)
        else:
            persona_seleccionada = st.session_state.usuario_actual
            st.caption(f"Mostrando datos de: **{persona_seleccionada}**")

        horas_excepcion = calcular_excepciones_mes(persona_seleccionada, anio, mes, st.session_state.excepciones)
        capacidad_real_persona = capacidad_base_mes - horas_excepcion
        df_persona = df_melt[df_melt['Operario'] == persona_seleccionada]

        if df_persona.empty and horas_excepcion == 0:
            st.info(f"{persona_seleccionada} no tiene horas cargadas en {MESES_ES[mes]} {anio}")
        else:
            ocupacion_persona = df_persona.groupby('Tarea')['Horas'].sum().reset_index() if not df_persona.empty else pd.DataFrame({'Tarea': [], 'Horas': []})
            total_persona = ocupacion_persona['Horas'].sum() if not ocupacion_persona.empty else 0

            # Nuevo cálculo: horas "Disponible" vs Trabajadas
            horas_disponible = df_persona[df_persona['Tarea'].isin(TAREAS_DISPONIBLE)]['Horas'].sum()
            horas_trabajadas = total_persona - horas_disponible
            porcentaje_disponible = (horas_disponible / total_persona * 100) if total_persona > 0 else 0

            # Semáforo nuevo: 0-60 verde, 60-80 amarillo, 80+ rojo
            if porcentaje_disponible > 80:
                color_semaforo = "🔴 Al límite"
            elif porcentaje_disponible >= 60:
                color_semaforo = "🟡 Atención"
            else:
                color_semaforo = "🟢 OK"

            col_torta, col_detalle = st.columns([2,1])
            with col_torta:
                if not ocupacion_persona.empty:
                    fig_persona = px.pie(ocupacion_persona, values='Horas', names='Tarea',
                                         color='Tarea', color_discrete_map=COLORES_TAREAS)
                    fig_persona.update_traces(textposition='inside', textinfo='percent+label')
                    fig_persona.update_layout(title=f"{total_persona:.1f}hs cargadas")
                    st.plotly_chart(fig_persona, use_container_width=True)
                else:
                    st.info("Sin horas cargadas este mes")

            with col_detalle:
                st.metric("Total cargado", f"{total_persona:.1f} hs")
                st.metric("Horas Trabajadas", f"{horas_trabajadas:.1f} hs")
                st.metric("Tiempo Disponible", f"{horas_disponible:.1f} hs",
                         delta=f"{porcentaje_disponible:.1f}% del total")
                st.metric("Capacidad real", f"{capacidad_real_persona:.0f} hs",
                         delta=f"-{horas_excepcion:.0f}hs excepción" if horas_excepcion > 0 else None)
                st.metric("Estado", color_semaforo)

                # Detalle de qué compone el "Disponible"
                if horas_disponible > 0:
                    with st.expander("Ver composición del Tiempo Disponible"):
                        detalle_disp = df_persona[df_persona['Tarea'].isin(TAREAS_DISPONIBLE)].groupby('Tarea')['Horas'].sum().reset_index()
                        detalle_disp.columns = ['Tarea', 'Horas']
                        detalle_disp['% del Total'] = (detalle_disp['Horas'] / total_persona * 100).round(1)
                        st.dataframe(detalle_disp, hide_index=True, use_container_width=True)

                pdf_individual = generar_pdf_reporte(
                    "individual", persona_seleccionada, mes, anio,
                    capacidad_real_persona, total_persona, porcentaje_disponible, color_semaforo, ocupacion_persona
                )
                st.download_button(
                    label="📄 Descargar Reporte PDF",
                    data=pdf_individual,
                    file_name=f"reporte_{persona_seleccionada}_{MESES_ES[mes]}_{anio}.pdf",
                    mime="application/pdf"
                )

        st.subheader("Histórico Últimos 3 Meses")
        df_historico = st.session_state.cargas.copy()
        df_historico['Fecha'] = pd.to_datetime(df_historico['Fecha'], errors='coerce')
        meses_historico = []
        for i in range(3):
            mes_calc = mes - i
            anio_calc = anio
            if mes_calc <= 0:
                mes_calc += 12
                anio_calc -= 1
            meses_historico.append((anio_calc, mes_calc))

        data_historico = []
        for a, m in reversed(meses_historico):
            mask = (df_historico['Fecha'].dt.month == m) & (df_historico['Fecha'].dt.year == a)
            df_mes_hist = df_historico[mask]
            total_mes = df_mes_hist[persona_seleccionada].sum()
            disponible_mes = df_mes_hist[df_mes_hist['Tarea'].isin(TAREAS_DISPONIBLE)][persona_seleccionada].sum()
            porcentaje_disp_mes = (disponible_mes / total_mes * 100) if total_mes > 0 else 0

            data_historico.append({
                'Mes': f"{MESES_ES[m][:3]} {a}",
                'Horas Totales': total_mes,
                '% Disponible': porcentaje_disp_mes
            })

        df_hist = pd.DataFrame(data_historico)
        fig_hist = go.Figure()
        fig_hist.add_trace(go.Bar(x=df_hist['Mes'], y=df_hist['Horas Totales'], name='Horas Totales', marker_color='#00BFFF'))
        fig_hist.add_trace(go.Scatter(x=df_hist['Mes'], y=df_hist['% Disponible'], name='% Disponible',
                                     line=dict(color='orange', dash='dash'), yaxis='y2'))

        fig_hist.update_layout(
            title=f"Evolución de {persona_seleccionada}",
            yaxis=dict(title="Horas Totales"),
            yaxis2=dict(title="% Disponible", overlaying='y', side='right', range=[0, 100])
        )
        st.plotly_chart(fig_hist, use_container_width=True)

        if st.session_state.usuario_actual == "Admin - Ver todo":
            st.divider()
            st.subheader("Distribución del Equipo Total")
            distribucion = df_melt.groupby('Tarea')['Horas'].sum().reset_index()
            total_equipo = distribucion['Horas'].sum()

            resumen_equipo = []
            for op in OPERARIOS_FIJOS:
                df_op = df_melt[df_melt['Operario'] == op]
                horas_op = df_op['Horas'].sum()
                if horas_op > 0:
                    disponible_op = df_op[df_op['Tarea'].isin(TAREAS_DISPONIBLE)]['Horas'].sum()
                    porcentaje = (disponible_op / horas_op * 100)
                    if porcentaje > 80:
                        estado = "🔴 Al límite"
                    elif porcentaje >= 60:
                        estado = "🟡 Atención"
                    else:
                        estado = "🟢 OK"

                    exc_op = calcular_excepciones_mes(op, anio, mes, st.session_state.excepciones)
                    cap_real_op = capacidad_base_mes - exc_op

                    resumen_equipo.append({
                        'Operario': op,
                        'Total Hs': f"{horas_op:.1f}",
                        'Disponible': f"{disponible_op:.1f}",
                        '% Disponible': f"{porcentaje:.1f}%",
                        'Estado': estado
                    })

            col_torta_eq, col_detalle_eq = st.columns([2,1])
            with col_torta_eq:
                fig_estudio = px.pie(distribucion, values='Horas', names='Tarea',
                             color='Tarea', color_discrete_map=COLORES_TAREAS)
                fig_estudio.update_traces(textposition='inside', textinfo='percent+label')
                fig_estudio.update_layout(title=f"Total equipo: {total_equipo:.1f}hs")
                st.plotly_chart(fig_estudio, use_container_width=True)

            with col_detalle_eq:
                st.metric("Total Equipo", f"{total_equipo:.1f} hs")
                disponible_equipo = df_melt[df_melt['Tarea'].isin(TAREAS_DISPONIBLE)]['Horas'].sum()
                porc_disp_equipo = (disponible_equipo / total_equipo * 100) if total_equipo > 0 else 0
                st.metric("Disponible Equipo", f"{disponible_equipo:.1f} hs", f"{porc_disp_equipo:.1f}%")

            st.subheader("Resumen del Equipo")
            df_resumen = pd.DataFrame(resumen_equipo)
            st.dataframe(df_resumen, use_container_width=True, hide_index=True)

            pdf_equipo = generar_pdf_reporte(
                "equipo", "Equipo", mes, anio,
                0, total_equipo, porc_disp_equipo, "", distribucion, df_resumen
            )
            st.download_button(
                label="📄 Descargar Reporte Equipo PDF",
                data=pdf_equipo,
                file_name=f"reporte_equipo_{MESES_ES[mes]}_{anio}.pdf",
                mime="application/pdf"
            )

# ===== CARGAR HORAS =====
elif menu in ["Cargar Mis Horas", "Cargar Horas"]:
    st.title("Cargar Horas")
    if st.session_state.usuario_actual == "Admin - Ver todo":
        usuario_carga = st.selectbox("Cargar horas para:", OPERARIOS_FIJOS)
    else:
        usuario_carga = st.session_state.usuario_actual
        st.info(f"Cargando horas para: **{usuario_carga}**")
    placeholder_exito = st.empty()
    with st.form("form_carga", clear_on_submit=True):
        st.markdown("**Formato: Día/Mes/Año** - Ej: 07/04/2026 es 7 de abril")
        fecha = st.date_input("Fecha", value=datetime.now(), format="DD/MM/YYYY")
        if fecha:
            dia_semana = DIAS_SEMANA_ES[fecha.weekday()]
            st.success(f"📅 Confirmado: **{dia_semana} {fecha.day} de {MESES_ES[fecha.month]} {fecha.year}**")
        tarea = st.selectbox("Área/Tarea", list(COLORES_TAREAS.keys()))
        horas = st.number_input(f"Horas trabajadas por {usuario_carga}", min_value=0.0, value=0.0, step=0.5)
        nota = st.text_area("Nota - opcional", placeholder="Ej: Cliente López, cierre mensual")
        submitted = st.form_submit_button("Guardar Carga")
        if submitted and horas > 0:
            nueva_fila = {'Fecha': fecha, 'Tarea': tarea, 'Nota': nota}
            for op in OPERARIOS_FIJOS:
                nueva_fila[op] = horas if op == usuario_carga else 0
            nueva_carga = pd.DataFrame([nueva_fila])
            st.session_state.cargas = pd.concat([st.session_state.cargas, nueva_carga], ignore_index=True)
            guardar_df("Cargas", st.session_state.cargas)
            with placeholder_exito:
                st.success(f"✅ **GUARDADO CORRECTAMENTE** - {horas}hs de {tarea} para {usuario_carga} el {fecha.strftime('%d/%m/%Y')}")
            time.sleep(3)
            st.rerun()
    st.subheader("Mis Cargas Registradas")
    col_filtro1, col_filtro2 = st.columns([1, 3])
    with col_filtro1:
        ver_solo_dia = st.checkbox("Solo este día", value=False)
    df_mis_cargas = st.session_state.cargas.copy()
    df_mis_cargas = df_mis_cargas[df_mis_cargas[usuario_carga] > 0]
    df_mis_cargas['Fecha'] = pd.to_datetime(df_mis_cargas['Fecha'], errors='coerce')
    if ver_solo_dia and fecha:
        df_mis_cargas = df_mis_cargas[df_mis_cargas['Fecha'].dt.date == fecha]
        if df_mis_cargas.empty:
            st.info(f"No tenés cargas registradas el {fecha.strftime('%d/%m/%Y')}")
    df_mis_cargas = df_mis_cargas.sort_values('Fecha', ascending=False)
    if df_mis_cargas.empty and not ver_solo_dia:
        st.info("No tenés cargas registradas")
    elif not df_mis_cargas.empty:
        if ver_solo_dia:
            total_dia = df_mis_cargas[usuario_carga].sum()
            st.metric(f"Total cargado el {fecha.strftime('%d/%m/%Y')}", f"{total_dia:.1f} hs")
        for i, row in df_mis_cargas.iterrows():
            col1, col2 = st.columns([5,1])
            with col1:
                if pd.notna(row['Fecha']):
                    fecha_dt = pd.to_datetime(row['Fecha'], errors='coerce')
                    if pd.notna(fecha_dt):
                        fecha_formateada = fecha_dt.strftime('%d/%m/%Y')
                        dia_sem = DIAS_SEMANA_ES[fecha_dt.weekday()]
                        st.write(f"**{fecha_formateada} ({dia_sem}) - {row['Tarea']}** | {row[usuario_carga]}hs | {row['Nota']}")
                    else:
                        st.write(f"**Fecha inválida - {row['Tarea']}** | {row[usuario_carga]}hs | {row['Nota']}")
                else:
                    st.write(f"**Sin fecha - {row['Tarea']}** | {row[usuario_carga]}hs | {row['Nota']}")
            with col2:
                if st.button("Eliminar", key=f"del_{i}"):
                    st.session_state.cargas = st.session_state.cargas.drop(i).reset_index(drop=True)
                    guardar_df("Cargas", st.session_state.cargas)
                    st.success("✅ Eliminado")
                    st.rerun()

# ===== CARGA MASIVA =====
elif menu == "Carga Masiva":
    st.title("Carga Masiva de Horas")
    st.caption("Usá esto para cargar meses completos atrasados sin volverte loco")
    if st.session_state.usuario_actual!= "Admin - Ver todo":
        st.warning("Solo el Admin puede usar carga masiva")
        st.stop()
    with st.form("form_masiva", clear_on_submit=False):
        col1, col2 = st.columns(2)
        with col1:
            usuario_masivo = st.selectbox("Operario", OPERARIOS_FIJOS)
            fecha_inicio = st.date_input("Fecha inicio", value=datetime(2026, 2, 1), format="DD/MM/YYYY")
        with col2:
            tarea_masiva = st.selectbox("Tarea/Área", list(COLORES_TAREAS.keys()))
            fecha_fin = st.date_input("Fecha fin", value=datetime(2026, 2, 28), format="DD/MM/YYYY")
        st.divider()
        col3, col4 = st.columns(2)
        with col3:
            horas_totales = st.number_input("Total de horas del período", min_value=0.0, value=0.0, step=1.0,
                                            help="Ej: 108hs para todo febrero")
        with col4:
            solo_habiles = st.checkbox("Solo días hábiles (lun-vie)", value=True,
                                       help="Tildado: reparte solo L-V. Destildado: reparte todos los días")
        nota_masiva = st.text_input("Nota para todos los registros", value="Carga masiva",
                                    placeholder="Ej: Total Febrero 2026")
        submitted = st.form_submit_button("📊 Previsualizar distribución")
        if submitted and horas_totales > 0 and fecha_inicio <= fecha_fin:
            if solo_habiles:
                dias_lista = pd.bdate_range(start=fecha_inicio, end=fecha_fin, freq='C', holidays=FERIADOS)
                tipo_dias = "días hábiles"
            else:
                dias_lista = pd.date_range(start=fecha_inicio, end=fecha_fin)
                tipo_dias = "días totales"
            cant_dias = len(dias_lista)
            horas_por_dia = horas_totales / cant_dias if cant_dias > 0 else 0
            st.session_state.preview_masiva = {
                'dias': dias_lista, 'horas_por_dia': horas_por_dia, 'cant_dias': cant_dias,
                'tipo_dias': tipo_dias, 'usuario': usuario_masivo, 'tarea': tarea_masiva, 'nota': nota_masiva
            }
    if 'preview_masiva' in st.session_state:
        prev = st.session_state.preview_masiva
        st.divider()
        st.subheader("Previsualización")
        col_m1, col_m2, col_m3 = st.columns(3)
        with col_m1:
            st.metric("Días a generar", f"{prev['cant_dias']} {prev['tipo_dias']}")
        with col_m2:
            st.metric("Horas por día", f"{prev['horas_por_dia']:.2f} hs")
        with col_m3:
            st.metric("Total", f"{prev['horas_por_dia'] * prev['cant_dias']:.1f} hs")
        st.info(f"Se van a crear {prev['cant_dias']} registros para **{prev['usuario']}** del {prev['dias'][0].strftime('%d/%m/%Y')} al {prev['dias'][-1].strftime('%d/%m/%Y')}")
        with st.expander("Ver primeras 5 filas de ejemplo"):
            ejemplo = pd.DataFrame({
                'Fecha': [d.strftime('%d/%m/%Y') for d in prev['dias'][:5]],
                'Tarea': [prev['tarea']] * min(5, prev['cant_dias']),
                prev['usuario']: [f"{prev['horas_por_dia']:.2f}"] * min(5, prev['cant_dias']),
                'Nota': [prev['nota']] * min(5, prev['cant_dias'])
            })
            st.dataframe(ejemplo, hide_index=True)
        col_btn1, col_btn2 = st.columns(2)
        with col_btn1:
            if st.button("✅ Confirmar y Guardar Todo", type="primary", use_container_width=True):
                nuevas_filas = []
                for dia in prev['dias']:
                    nueva_fila = {'Fecha': dia.date(), 'Tarea': prev['tarea'], 'Nota': prev['nota']}
                    for op in OPERARIOS_FIJOS:
                        nueva_fila[op] = round(prev['horas_por_dia'], 2) if op == prev['usuario'] else 0
                    nuevas_filas.append(nueva_fila)
                df_nuevo = pd.DataFrame(nuevas_filas)
                st.session_state.cargas = pd.concat([st.session_state.cargas, df_nuevo], ignore_index=True)
                guardar_df("Cargas", st.session_state.cargas)
                st.success(f"✅ **{prev['cant_dias']} registros creados** para {prev['usuario']}. Total: {prev['horas_por_dia'] * prev['cant_dias']:.1f}hs")
                st.balloons()
                del st.session_state.preview_masiva
                time.sleep(2)
                st.rerun()
        with col_btn2:
            if st.button("❌ Cancelar", use_container_width=True):
                del st.session_state.preview_masiva
                st.rerun()

# ===== EXCEPCIONES =====
elif menu == "Excepciones":
    st.title("Excepciones - Vacaciones/Licencias")
    with st.form("form_excepcion", clear_on_submit=True):
        col1, col2 = st.columns(2)
        with col1:
            operario = st.selectbox("Operario", OPERARIOS_FIJOS)
            fecha_inicio_exc = st.date_input("Fecha Inicio", format="DD/MM/YYYY")
        with col2:
            fecha_fin_exc = st.date_input("Fecha Fin", format="DD/MM/YYYY")
            motivo = st.selectbox("Motivo", ["Vacaciones", "Licencia", "Enfermedad", "Otro"])
        if st.form_submit_button("Guardar Excepción"):
            dias = calcular_dias_habiles(fecha_inicio_exc, fecha_fin_exc)
            horas = dias * HORAS_DIA_LABORAL
            nuevo = pd.DataFrame([{
                'Operario': operario, 'Fecha Inicio': fecha_inicio_exc,
                'Fecha Fin': fecha_fin_exc, 'Motivo': motivo, 'Horas': horas
            }])
            st.session_state.excepciones = pd.concat([st.session_state.excepciones, nuevo], ignore_index=True)
            if guardar_df("Excepciones", st.session_state.excepciones):
                st.success(f"✅ Excepción guardada: {operario} - {dias} días = {horas}hs menos")
                time.sleep(1)
                st.rerun()
            else:
                st.error("❌ No se pudo guardar. Revisá la hoja Excepciones en Sheets.")
    st.subheader("Excepciones Cargadas")
    if st.session_state.excepciones.empty:
        st.info("No hay excepciones cargadas")
    else:
        st.dataframe(st.session_state.excepciones, use_container_width=True, hide_index=True)

# ===== EXPORTAR EXCEL =====
elif menu == "Exportar Excel":
    st.title("Exportar a Excel")
    if st.session_state.cargas.empty:
        st.warning("No hay datos para exportar")
    else:
        csv_cargas = st.session_state.cargas.to_csv(index=False).encode('utf-8')
        csv_exc = st.session_state.excepciones.to_csv(index=False).encode('utf-8')
        col1, col2 = st.columns(2)
        col1.download_button("⬇️ Descargar Cargas.csv", csv_cargas, "cargas.csv", "text/csv")
        col2.download_button("⬇️ Descargar Excepciones.csv", csv_exc, "excepciones.csv", "text/csv")

# ===== RESETEAR DATOS - SOLO ADMIN =====
elif menu == "Resetear Datos":
    st.title("⚠️ Resetear Datos")
    st.error("Esto borra TODAS las cargas del sistema. No se puede deshacer.")
    st.info("Las cargas viejas de Natalia y Athina que viste son datos de prueba. Usá esto para empezar limpio.")
    confirmacion = st.text_input("Escribí BORRAR para confirmar", placeholder="BORRAR")
    if confirmacion == "BORRAR":
        if st.button("🗑️ Borrar todas las cargas", type="primary"):
            df_vacio = pd.DataFrame(columns=['Fecha', 'Tarea'] + OPERARIOS_FIJOS + ['Nota'])
            guardar_df("Cargas", df_vacio)
            st.session_state.cargas = df_vacio
            st.success("✅ Base limpiada. Todas las cargas fueron eliminadas.")
            st.balloons()
            st.rerun()

