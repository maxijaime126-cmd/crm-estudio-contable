import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import plotly.express as px
import plotly.graph_objects as go
import gspread
from google.oauth2.service_account import Credentials

st.set_page_config(page_title="CRM Capacidad Instalada", layout="wide")

# ===== CONFIGURACIÓN =====
COLORES_TAREAS = {
    "DOCUMENTACIÓN (incluye carga y control)": "#FFB6C1",
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
                df[col] = pd.to_datetime(df[col], dayfirst=True, errors='coerce').dt.date
        return df, ws
    except:
        return pd.DataFrame(), None

def guardar_df(nombre_hoja, df):
    ws = conectar_sheets().worksheet(nombre_hoja)
    ws.clear()

    # FIX: Convertir fechas y NaN para que gspread no rompa
    df_copy = df.copy()
    for col in df_copy.columns:
        if pd.api.types.is_datetime64_any_dtype(df_copy[col]):
            df_copy[col] = df_copy[col].dt.strftime('%Y-%m-%d')

    df_copy = df_copy.fillna('').astype(str)

    ws.update([df_copy.columns.values.tolist()] + df_copy.values.tolist())
    st.cache_data.clear()

# Cargar feriados
feriados_df, _ = cargar_hoja("Feriados")
if not feriados_df.empty and 'fecha' in feriados_df.columns:
    FERIADOS = pd.to_datetime(feriados_df['fecha']).dt.date.tolist()
else:
    FERIADOS = []

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
    menu = st.sidebar.radio("Menú", ["Panel de Control", "Cargar Horas", "Excepciones", "Exportar Excel"])
else:
    menu = st.sidebar.radio("Menú", ["Panel de Control", "Cargar Mis Horas"])

# ===== PANEL DE CONTROL =====
if menu == "Panel de Control":
    st.title("Panel de Control - Capacidad Instalada")

    col1, col2 = st.columns(2)
    with col1:
        anio = st.selectbox("Año", [2025, 2026, 2027], index=1)
    with col2:
        mes = st.selectbox("Mes", list(range(1,13)),
                           format_func=lambda x: datetime(1900, x, 1).strftime('%B'),
                           index=3)

    dias_habiles, capacidad_base = calcular_capacidad_mensual(anio, mes)
    st.write(f"**{datetime(anio, mes, 1).strftime('%B %Y')} - {dias_habiles} días hábiles | Capacidad base: {capacidad_base}hs por persona**")

    # Filtrar cargas del mes seleccionado
    df_mes = st.session_state.cargas.copy()
    if not df_mes.empty:
        df_mes['Fecha'] = pd.to_datetime(df_mes['Fecha'], errors='coerce')
        mask = (df_mes['Fecha'].dt.month == mes) & (df_mes['Fecha'].dt.year == anio)
        df_mes = df_mes[mask]

    if df_mes.empty:
        st.info("Cargá horas para ver el gráfico")
    else:
        # Preparar datos
        df_melt = df_mes.melt(id_vars=['Fecha', 'Tarea', 'Nota'],
                              value_vars=OPERARIOS_FIJOS,
                              var_name='Operario', value_name='Horas')
        df_melt = df_melt[df_melt['Horas'] > 0]

        # ===== GRÁFICO 1: TORTA INDIVIDUAL CON FILTRO =====
        st.subheader("Ocupación Individual")
        if st.session_state.usuario_actual == "Admin - Ver todo":
            persona_seleccionada = st.selectbox("Seleccionar persona", OPERARIOS_FIJOS)
        else:
            persona_seleccionada = st.session_state.usuario_actual
            st.caption(f"Mostrando datos de: **{persona_seleccionada}**")

        df_persona = df_melt[df_melt['Operario'] == persona_seleccionada]
        if df_persona.empty:
            st.info(f"{persona_seleccionada} no tiene horas cargadas en {datetime(anio, mes, 1).strftime('%B %Y')}")
        else:
            ocupacion_persona = df_persona.groupby('Tarea')['Horas'].sum().reset_index()
            total_persona = ocupacion_persona['Horas'].sum()

            porcentaje = (total_persona / capacidad_base * 100)
            if porcentaje > 100:
                color_semaforo = "🔴 Sobrecarga"
            elif porcentaje >= 80:
                color_semaforo = "🟡 Al límite"
            else:
                color_semaforo = "🟢 OK"

            col_torta, col_detalle = st.columns([2,1])
            with col_torta:
                fig_persona = px.pie(ocupacion_persona, values='Horas', names='Tarea',
                                     color='Tarea', color_discrete_map=COLORES_TAREAS)
                fig_persona.update_traces(textposition='inside', textinfo='percent+label')
                fig_persona.update_layout(title=f"{total_persona:.1f}hs de {capacidad_base}hs")
                st.plotly_chart(fig_persona, use_container_width=True)

            with col_detalle:
                st.metric("Total cargado", f"{total_persona:.1f} hs")
                st.metric("Capacidad", f"{capacidad_base} hs")
                st.metric("Ocupación", f"{porcentaje:.1f}%", delta=f"{color_semaforo}")

        # ===== HISTÓRICO 3 MESES =====
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
            total_mes = df_historico[mask][persona_seleccionada].sum()
            data_historico.append({
                'Mes': datetime(a, m, 1).strftime('%b %Y'),
                'Horas': total_mes,
                'Capacidad': calcular_capacidad_mensual(a, m)[1]
            })

        df_hist = pd.DataFrame(data_historico)
        fig_hist = go.Figure()
        fig_hist.add_trace(go.Bar(x=df_hist['Mes'], y=df_hist['Horas'], name='Horas Cargadas', marker_color='#00BFFF'))
        fig_hist.add_trace(go.Scatter(x=df_hist['Mes'], y=df_hist['Capacidad'], name='Capacidad',
                                     line=dict(color='red', dash='dash')))
        fig_hist.update_layout(title=f"Evolución de {persona_seleccionada}", yaxis_title="Horas")
        st.plotly_chart(fig_hist, use_container_width=True)

        # ===== GRÁFICO 2: TORTA DEL EQUIPO TOTAL - SOLO ADMIN =====
        if st.session_state.usuario_actual == "Admin - Ver todo":
            st.divider()
            st.subheader("Distribución del Equipo Total")

            distribucion = df_melt.groupby('Tarea')['Horas'].sum().reset_index()
            total_equipo = distribucion['Horas'].sum()

            col_torta_eq, col_detalle_eq = st.columns([2,1])
            with col_torta_eq:
                fig_estudio = px.pie(distribucion, values='Horas', names='Tarea',
                             color='Tarea', color_discrete_map=COLORES_TAREAS)
                fig_estudio.update_traces(textposition='inside', textinfo='percent+label')
                fig_estudio.update_layout(title=f"Total equipo: {total_equipo:.1f}hs")
                st.plotly_chart(fig_estudio, use_container_width=True)

            with col_detalle_eq:
                st.metric("Total Equipo", f"{total_equipo:.1f} hs")
                st.metric("Capacidad Total", f"{capacidad_base * 4:.1f} hs")
                ocup_equipo = (total_equipo / (capacidad_base * 4) * 100)
                st.metric("Ocupación Equipo", f"{ocup_equipo:.1f}%")

            st.subheader("Resumen del Equipo")
            resumen = df_melt.groupby('Operario')['Horas'].sum().reset_index()
            resumen['Capacidad'] = capacidad_base
            resumen['% Ocupación'] = (resumen['Horas'] / capacidad_base * 100).round(1)
            resumen['Estado'] = resumen['% Ocupación'].apply(lambda x: "🔴" if x > 100 else "🟡" if x >= 80 else "🟢")
            resumen['% Ocupación'] = resumen['% Ocupación'].astype(str) + '%'
            st.dataframe(resumen, use_container_width=True, hide_index=True)

# ===== CARGAR HORAS =====
elif menu in ["Cargar Mis Horas", "Cargar Horas"]:
    st.title("Cargar Horas")

    if st.session_state.usuario_actual == "Admin - Ver todo":
        usuario_carga = st.selectbox("Cargar horas para:", OPERARIOS_FIJOS)
    else:
        usuario_carga = st.session_state.usuario_actual
        st.info(f"Cargando horas para: **{usuario_carga}**")

    with st.form("form_carga", clear_on_submit=True):
        fecha = st.date_input("Fecha", value=datetime.now())
        tarea = st.selectbox("Área/Tarea", list(COLORES_TAREAS.keys()))
        horas = st.number_input(f"Horas trabajadas por {usuario_carga}", min_value=0.0, value=0.0, step=0.5)
        nota = st.text_area("Nota - opcional", placeholder="Ej: Cliente López, cierre mensual")

        if st.form_submit_button("Guardar Carga") and horas > 0:
            nueva_fila = {'Fecha': fecha, 'Tarea': tarea, 'Nota': nota}
            for op in OPERARIOS_FIJOS:
                nueva_fila[op] = horas if op == usuario_carga else 0

            nueva_carga = pd.DataFrame([nueva_fila])
            st.session_state.cargas = pd.concat([st.session_state.cargas, nueva_carga], ignore_index=True)
            guardar_df("Cargas", st.session_state.cargas)
            st.success(f"✅ Carga guardada: {horas}hs de {tarea} para {usuario_carga}")
            st.rerun()

    st.subheader("Mis Cargas Registradas")
    df_mis_cargas = st.session_state.cargas.copy()
    df_mis_cargas = df_mis_cargas[df_mis_cargas[usuario_carga] > 0]

    if df_mis_cargas.empty:
        st.info("No tenés cargas registradas")
    else:
        for i, row in df_mis_cargas.iterrows():
            col1, col2 = st.columns([5,1])
            with col1:
                st.write(f"**{row['Fecha']} - {row['Tarea']}** | {row[usuario_carga]}hs | {row['Nota']}")
            with col2:
                if st.button("Eliminar", key=f"del_{i}"):
                    st.session_state.cargas = st.session_state.cargas.drop(i).reset_index(drop=True)
                    guardar_df("Cargas", st.session_state.cargas)
                    st.success("✅ Eliminado")
                    st.rerun()

# ===== EXCEPCIONES =====
elif menu == "Excepciones":
    st.title("Excepciones - Vacaciones/Licencias")

    with st.form("form_excepcion", clear_on_submit=True):
        col1, col2 = st.columns(2)
        with col1:
            operario = st.selectbox("Operario", OPERARIOS_FIJOS)
            fecha_inicio_exc = st.date_input("Fecha Inicio")
        with col2:
            fecha_fin_exc = st.date_input("Fecha Fin")
            motivo = st.selectbox("Motivo", ["Vacaciones", "Licencia", "Enfermedad", "Otro"])

        if st.form_submit_button("Guardar Excepción"):
            dias = calcular_dias_habiles(fecha_inicio_exc, fecha_fin_exc)
            horas = dias * HORAS_DIA_LABORAL
            nuevo = pd.DataFrame([{
                'Operario': operario, 'Fecha Inicio': fecha_inicio_exc,
                'Fecha Fin': fecha_fin_exc, 'Motivo': motivo, 'Horas': horas
            }])
            st.session_state.excepciones = pd.concat([st.session_state.excepciones, nuevo], ignore_index=True)
            guardar_df("Excepciones", st.session_state.excepciones)
            st.success(f"✅ Excepción: {operario} - {dias} días = {horas}hs menos")
            st.rerun()

    st.subheader("Excepciones Cargadas")
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
