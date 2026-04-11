# rebuild python 3.11
import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import plotly.express as px
import plotly.graph_objects as go
import gspread
from oauth2client.service_account import ServiceAccountCredentials

st.set_page_config(page_title="CRM Capacidad Instalada", layout="wide")

# CONFIGURACIÓN - TAREAS Y COLORES
COLORES_TAREAS = {
    "DOCUMENTACIÓN (incluye carga y control)": "#FFB6C1", # Rosa
    "IMPUESTOS": "#FF00FF", # Fucsia
    "SUELDOS": "#FFFF00", # Amarillo
    "CONTABILIDAD": "#00FF00", # Verde
    "ATENCION AL CLIENTE": "#00BFFF", # Celeste
    "TAREAS NO RUTINARIAS": "#ADD8E6", # Celeste claro
    "DISPONIBLE": "#FFFFFF", # Blanco
    "REUNIONES DE EQUIPO": "#E6E6FA", # Lila
    "INASISTENCIA POR EXAMEN O TRAMITE": "#FFDAB9", # Durazno
    "PLANIFICACIONES/ORGANIZACIÓN/PROCEDIMIENTO S/INFORMES": "#40E0D0" # Turquesa
}

OPERARIOS_FIJOS = ["Natalia", "Maximiliano", "Athina", "Johana"]
HORAS_DIA_LABORAL = 6

# PEGÁ ACÁ LA URL DE TU GOOGLE SHEET ENTRE LAS COMILLAS
URL_GOOGLE_SHEET = "https://docs.google.com/spreadsheets/d/1Y6T-GcFWQHWVWkvbNYRgQg9tzTU3moc6ms-o23JpwXE/edit"

# CONEXIÓN A GOOGLE SHEETS
@st.cache_resource
def conectar_sheets():
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_dict(st.secrets["gcp_service_account"], scope)
    client = gspread.authorize(creds)
    sheet = client.open_by_url(URL_GOOGLE_SHEET)
    return sheet

def cargar_cargas():
    sheet = conectar_sheets()
    worksheet = sheet.worksheet("Hoja 1")
    data = worksheet.get_all_records()
    df = pd.DataFrame(data)
    if not df.empty:
        df['Fecha'] = pd.to_datetime(df['Fecha'], dayfirst=True).dt.date
    return df

def guardar_carga(nueva_fila):
    sheet = conectar_sheets()
    worksheet = sheet.worksheet("Hoja 1")
    worksheet.append_row([
        nueva_fila['Fecha'].strftime('%d/%m/%Y'),
        nueva_fila['Tarea'],
        nueva_fila['Natalia'],
        nueva_fila['Maximiliano'],
        nueva_fila['Athina'],
        nueva_fila['Johana'],
        nueva_fila['Nota']
    ])

def eliminar_carga(idx_real):
    sheet = conectar_sheets()
    worksheet = sheet.worksheet("Hoja 1")
    worksheet.delete_rows(idx_real + 2) # +2 porque sheets empieza en 1 y fila 1 son títulos

def cargar_excepciones():
    sheet = conectar_sheets()
    try:
        worksheet = sheet.worksheet("Excepciones")
        data = worksheet.get_all_records()
        df = pd.DataFrame(data)
        if not df.empty:
            df['Fecha Inicio'] = pd.to_datetime(df['Fecha Inicio'], dayfirst=True).dt.date
            df['Fecha Fin'] = pd.to_datetime(df['Fecha Fin'], dayfirst=True).dt.date
        return df
    except:
        return pd.DataFrame(columns=['Operario', 'Fecha Inicio', 'Fecha Fin', 'Motivo', 'Horas'])

def guardar_excepcion(nueva_fila):
    sheet = conectar_sheets()
    worksheet = sheet.worksheet("Excepciones")
    worksheet.append_row([
        nueva_fila['Operario'],
        nueva_fila['Fecha Inicio'].strftime('%d/%m/%Y'),
        nueva_fila['Fecha Fin'].strftime('%d/%m/%Y'),
        nueva_fila['Motivo'],
        nueva_fila['Horas']
    ])

# Cargar feriados
try:
    feriados_df = pd.read_csv("feriados_2026.csv")
    FERIADOS = pd.to_datetime(feriados_df['fecha']).dt.date.tolist()
except:
    FERIADOS = []
    st.sidebar.warning("No se encontró feriados_2026.csv")

# FUNCIONES
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

# LOGIN DE USUARIO
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

# MOSTRAR USUARIO LOGUEADO Y BOTÓN SALIR
col_user, col_logout = st.sidebar.columns([3,1])
with col_user:
    st.sidebar.success(f"Usuario: **{st.session_state.usuario_actual}**")
with col_logout:
    if st.sidebar.button("Salir"):
        st.session_state.usuario_actual = None
        st.rerun()

# MENÚ
if st.session_state.usuario_actual == "Admin - Ver todo":
    menu = st.sidebar.radio("Menú", ["Panel de Control", "Cargar Horas", "Excepciones", "Exportar Excel"])
else:
    menu = st.sidebar.radio("Menú", ["Panel de Control", "Cargar Mis Horas"])

# CARGAR DATOS
cargas_df = cargar_cargas()
excepciones_df = cargar_excepciones()

# PANEL DE CONTROL
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

    df_mes = cargas_df.copy()
    if not df_mes.empty:
        df_mes['Fecha'] = pd.to_datetime(df_mes['Fecha'])
        mask = (df_mes['Fecha'].dt.month == mes) & (df_mes['Fecha'].dt.year == anio)
        df_mes = df_mes[mask]

    if df_mes.empty:
        st.info("Cargá horas para ver el gráfico")
    else:
        df_melt = df_mes.melt(id_vars=['Fecha', 'Tarea', 'Nota'],
                              value_vars=OPERARIOS_FIJOS,
                              var_name='Operario', value_name='Horas')
        df_melt = df_melt[df_melt['Horas'] > 0]

        st.subheader("Mi Ocupación Mensual")
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

        st.subheader("Histórico Últimos 3 Meses")
        df_historico = cargas_df.copy()
        df_historico['Fecha'] = pd.to_datetime(df_historico['Fecha'])
        
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

        if st.session_state.usuario_actual == "Admin - Ver todo":
            st.subheader("Distribución del Estudio")
            distribucion = df_melt.groupby('Tarea')['Horas'].sum().reset_index()
            
            fig_estudio = px.pie(distribucion, values='Horas', names='Tarea', 
                         color='Tarea', color_discrete_map=COLORES_TAREAS,
                         title='Total del estudio por área')
            fig_estudio.update_traces(textposition='inside', textinfo='percent+label')
            st.plotly_chart(fig_estudio, use_container_width=True)

            st.subheader("Resumen del Equipo")
            resumen = df_melt.groupby('Operario')['Horas'].sum().reset_index()
            resumen['Capacidad'] = capacidad_base
            resumen['% Ocupación'] = (resumen['Horas'] / capacidad_base * 100).round(1)
            resumen['Estado'] = resumen['% Ocupación'].apply(lambda x: "🔴" if x > 100 else "🟡" if x >= 80 else "🟢")
            resumen['% Ocupación'] = resumen['% Ocupación'].astype(str) + '%'
            st.dataframe(resumen, use_container_width=True)

# CARGAR HORAS
elif menu == "Cargar Mis Horas" or menu == "Cargar Horas":
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

        submitted = st.form_submit_button("Guardar Carga")
        if submitted and horas > 0:
            nueva_fila = {'Fecha': fecha, 'Tarea': tarea, 'Nota': nota}
            for op in OPERARIOS_FIJOS:
                nueva_fila[op] = horas if op == usuario_carga else 0
            
            guardar_carga(nueva_fila)
            st.success(f"✅ Carga guardada: {horas}hs de {tarea} para {usuario_carga}")
            st.cache_data.clear()
            st.rerun()

    st.subheader("Mis Cargas Registradas")
    df_mis_cargas = cargas_df[cargas_df[usuario_carga] > 0].copy()
    
    if df_mis_cargas.empty:
        st.info("No tenés cargas registradas")
    else:
        for i, row in df_mis_cargas.iterrows():
            col1, col2 = st.columns([6,1])
            with col1:
                st.write(f"**{row['Fecha']} - {row['Tarea']}** | {row[usuario_carga]}hs | {row['Nota']}")
            with col2:
                if st.button("Eliminar", key=f"del_{i}"):
                    eliminar_carga(i)
                    st.success("✅ Eliminado")
                    st.cache_data.clear()
                    st.rerun()

# EXCEPCIONES - Solo admin
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
            nuevo = {
                'Operario': operario, 'Fecha Inicio': fecha_inicio_exc,
                'Fecha Fin': fecha_fin_exc, 'Motivo': motivo, 'Horas': horas
            }
            guardar_excepcion(nuevo)
            st.success(f"✅ Excepción: {operario} - {dias} días = {horas}hs menos")
            st.cache_data.clear()
            st.rerun()

    st.subheader("Excepciones Cargadas")
    st.dataframe(excepciones_df, use_container_width=True)

# EXPORTAR EXCEL - Solo admin
elif menu == "Exportar Excel":
    st.title("Exportar a Excel")
    if cargas_df.empty:
        st.warning("No hay datos para exportar")
    else:
        csv = cargas_df.to_csv(index=False).encode('utf-8')
        st.download_button("Descargar Cargas CSV", csv, "cargas.csv", "text/csv")
