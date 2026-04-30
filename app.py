import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import plotly.express as px
import plotly.graph_objects as go
import gspread
from google.oauth2.service_account import Credentials
from io import BytesIO
import time
import os

# ===== 1. CONFIGURACIÓN DE PÁGINA =====
st.set_page_config(page_title="CRM Capacidad Instalada", layout="wide")

# ===== 2. CONFIGURACIÓN DE NEGOCIO =====
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
HORAS_DIA_LABORAL = 6[cite: 1]

TAREAS_DISPONIBLE_TIPO = [
    "DISPONIBLE",
    "PLANIFICACIONES/ORGANIZACIÓN/PROCEDIMIENTO S/INFORMES",
    "INASISTENCIA POR EXAMEN O TRAMITE"
]

MESES_ES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
    7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
}

# ===== 3. CARGA DE FERIADOS DESDE GITHUB =====
@st.cache_data
def cargar_feriados_desde_csv():
    try:
        df_f = pd.read_csv("feriados_2026.csv")
        df_f.columns = df_f.columns.str.strip().str.lower()
        if 'fecha' in df_f.columns:
            df_f['fecha'] = pd.to_datetime(df_f['fecha'], errors='coerce')[cite: 1]
            return df_f['fecha'].dt.date.dropna().tolist()
        return []
    except Exception:
        return []

FERIADOS = cargar_feriados_desde_csv()

# ===== 4. CONEXIÓN GOOGLE SHEETS =====
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
                df[col] = pd.to_datetime(df[col], dayfirst=True, errors='coerce')[cite: 1]
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
                df_copy[col] = pd.to_datetime(df_copy[col], errors='coerce').dt.strftime('%d/%m/%Y')[cite: 1]
        df_copy = df_copy.fillna('').astype(str)
        ws.update([df_copy.columns.values.tolist()] + df_copy.values.tolist())
        st.cache_data.clear()
        return True
    except Exception as e:
        st.error(f"Error guardando: {e}")
        return False

# ===== 5. FUNCIONES DE CAPACIDAD =====
def calcular_dias_habiles(fecha_inicio, fecha_fin):
    dias = pd.bdate_range(start=fecha_inicio, end=fecha_fin, freq='C', holidays=FERIADOS)[cite: 1]
    return len(dias)

def calcular_capacidad_mensual(anio, mes):
    inicio = datetime(anio, mes, 1)
    fin = (datetime(anio, mes + 1, 1) if mes < 12 else datetime(anio + 1, 1, 1)) - timedelta(days=1)
    dias_h = calcular_dias_habiles(inicio.date(), fin.date())
    return dias_h, dias_h * HORAS_DIA_LABORAL

def calcular_excepciones_mes(operario, anio, mes, df_excepciones):
    if df_excepciones.empty: return 0
    df_exc = df_excepciones.copy()
    df_exc = df_exc[df_exc['Operario'] == operario]
    total_hs = 0
    inicio_mes = datetime(anio, mes, 1).date()
    fin_mes = (datetime(anio, mes + 1, 1) if mes < 12 else datetime(anio + 1, 1, 1)).date() - timedelta(days=1)
    for _, row in df_exc.iterrows():
        if pd.notna(row['Fecha Inicio']) and pd.notna(row['Fecha Fin']):
            i_ef = max(row['Fecha Inicio'].date(), inicio_mes)
            f_ef = min(row['Fecha Fin'].date(), fin_mes)
            if i_ef <= f_ef:
                total_hs += calcular_dias_habiles(i_ef, f_ef) * HORAS_DIA_LABORAL
    return total_hs

# ===== 6. SESSION STATE & LOGIN =====
if 'cargas' not in st.session_state:
    df, _ = cargar_hoja("Cargas")
    st.session_state.cargas = df if not df.empty else pd.DataFrame(columns=['Fecha', 'Tarea'] + OPERARIOS_FIJOS + ['Nota'])

if 'excepciones' not in st.session_state:
    df, _ = cargar_hoja("Excepciones")
    st.session_state.excepciones = df if not df.empty else pd.DataFrame(columns=['Operario', 'Fecha Inicio', 'Fecha Fin', 'Motivo', 'Horas'])

if 'usuario_actual' not in st.session_state:
    st.session_state.usuario_actual = None

if st.session_state.usuario_actual is None:
    st.title("CRM Capacidad Instalada")
    usuario = st.selectbox("Usuario:", ["Seleccionar..."] + OPERARIOS_FIJOS + ["Admin - Ver todo"])
    if st.button("Ingresar") and usuario != "Seleccionar...":
        st.session_state.usuario_actual = usuario
        st.rerun()
    st.stop()

# ===== 7. SIDEBAR & MENU =====
st.sidebar.success(f"Usuario: {st.session_state.usuario_actual}")
if st.sidebar.button("Salir"):
    st.session_state.clear()
    st.rerun()

menu = st.sidebar.radio("Menú", ["Panel de Control", "Cargar Horas", "Excepciones", "Resetear Datos"] if st.session_state.usuario_actual == "Admin - Ver todo" else ["Panel de Control", "Cargar Mis Horas"])

# ===== 8. PANEL DE CONTROL =====
if menu == "Panel de Control":
    st.title("Panel de Control - Ocupación")
    
    col_a, col_m = st.columns(2)
    with col_a: anio = st.selectbox("Año", [2025, 2026, 2027], index=1)
    with col_m: mes = st.selectbox("Mes", list(range(1,13)), format_func=lambda x: MESES_ES[x], index=datetime.now().month - 1)

    dias_h, cap_base = calcular_capacidad_mensual(anio, mes)
    st.info(f"**{MESES_ES[mes]} {anio}**: {dias_h} días hábiles. Capacidad: {cap_base}hs.")

    df_pnl = st.session_state.cargas.copy()
    if not df_pnl.empty:
        df_pnl['Fecha'] = pd.to_datetime(df_pnl['Fecha'], errors='coerce')
        df_mes = df_pnl[(df_pnl['Fecha'].dt.month == mes) & (df_pnl['Fecha'].dt.year == anio)]

        if df_mes.empty:
            st.warning("No hay datos para este mes.")
        else:
            df_melt = df_mes.melt(id_vars=['Fecha', 'Tarea', 'Nota'], value_vars=OPERARIOS_FIJOS, var_name='Operario', value_name='Horas')
            df_melt['Horas'] = pd.to_numeric(df_melt['Horas'], errors='coerce').fillna(0).round(2)
            df_melt = df_melt[df_melt['Horas'] > 0]

            persona_sel = st.selectbox("Operario:", OPERARIOS_FIJOS) if st.session_state.usuario_actual == "Admin - Ver todo" else st.session_state.usuario_actual
            
            # --- SECCIÓN INDIVIDUAL ---
            hs_exc = calcular_excepciones_mes(persona_sel, anio, mes, st.session_state.excepciones)
            cap_real = cap_base - hs_exc
            df_ind = df_melt[df_melt['Operario'] == persona_sel]
            
            resumen_t = df_ind.groupby('Tarea')['Horas'].sum().round(1).reset_index()[cite: 1]
            total_c = resumen_t['Horas'].sum().round(1)
            
            hs_libres = df_ind[df_ind['Tarea'].isin(TAREAS_DISPONIBLE_TIPO)]['Horas'].sum().round(1)
            porc_libres = (hs_libres / total_c * 100) if total_c > 0 else 0

            # Semáforo Corregido[cite: 1]
            if porc_libres > 20: color_s = "🟢 OK"
            elif porc_libres >= 10: color_s = "🟡 Atención"
            else: color_s = "🔴 Al límite"

            c_graf, c_metr = st.columns([2,1])
            with c_graf:
                fig = px.pie(resumen_t, values='Horas', names='Tarea', color='Tarea', color_discrete_map=COLORES_TAREAS)
                fig.update_layout(title=f"Ocupación: {persona_sel} ({total_c} hs)")
                st.plotly_chart(fig, use_container_width=True)
            with c_metr:
                st.metric("Total Cargado", f"{total_c} hs")
                st.metric("Capacidad Real", f"{cap_real} hs")
                st.metric("Tiempo Disponible", f"{hs_libres} hs", delta=f"{porc_libres:.1f}%")
                st.subheader(f"Estado: {color_s}")

            # --- HISTÓRICO 3 MESES ---
            st.divider()
            st.subheader("Histórico Últimos 3 Meses")
            data_hist = []
            for i in range(3):
                m_hist = mes - i
                a_hist = anio
                if m_hist <= 0: m_hist += 12; a_hist -= 1
                df_m_h = df_pnl[(df_pnl['Fecha'].dt.month == m_hist) & (df_pnl['Fecha'].dt.year == a_hist)]
                t_m = df_m_h[persona_sel].sum().round(1)
                d_m = df_m_h[df_m_h['Tarea'].isin(TAREAS_DISPONIBLE_TIPO)][persona_sel].sum().round(1)
                p_d = (d_m / t_m * 100) if t_m > 0 else 0
                data_hist.append({'Mes': f"{MESES_ES[m_hist][:3]} {a_hist}", 'Horas': t_m, '% Disponible': p_d})
            
            df_h_plot = pd.DataFrame(data_hist[::-1])
            fig_h = go.Figure()
            fig_h.add_trace(go.Bar(x=df_h_plot['Mes'], y=df_h_plot['Horas'], name='Horas Totales', marker_color='#00BFFF'))
            fig_h.add_trace(go.Scatter(x=df_h_plot['Mes'], y=df_h_plot['% Disponible'], name='% Disponible', yaxis='y2', line=dict(color='orange', dash='dash')))
            fig_h.update_layout(yaxis2=dict(overlaying='y', side='right', range=[0, 100]), title="Evolución Personal")
            st.plotly_chart(fig_h, use_container_width=True)

            # --- SECCIÓN EQUIPO TOTAL (SOLO ADMIN) ---
            if st.session_state.usuario_actual == "Admin - Ver todo":
                st.divider()
                st.subheader("Distribución del Equipo Total")
                dist_eq = df_melt.groupby('Tarea')['Horas'].sum().round(1).reset_index()[cite: 1]
                total_eq = dist_eq['Horas'].sum().round(1)
                
                c_eq1, c_eq2 = st.columns([2,1])
                with c_eq1:
                    fig_eq = px.pie(dist_eq, values='Horas', names='Tarea', color='Tarea', color_discrete_map=COLORES_TAREAS)
                    fig_eq.update_layout(title=f"Estudio Completo ({total_eq} hs totales)")
                    st.plotly_chart(fig_eq, use_container_width=True)
                with c_eq2:
                    st.metric("Total Equipo", f"{total_eq} hs")
                    # Resumen por persona
                    res_op = []
                    for op in OPERARIOS_FIJOS:
                        h_op = df_melt[df_melt['Operario'] == op]['Horas'].sum().round(1)
                        res_op.append({'Operario': op, 'Horas': h_op})
                    st.table(pd.DataFrame(res_op))

# ===== 9. CARGAR HORAS =====
elif menu in ["Cargar Mis Horas", "Cargar Horas"]:
    st.title("Cargar Horas")
    u_c = st.selectbox("Persona:", OPERARIOS_FI_FIJOS) if st.session_state.usuario_actual == "Admin - Ver todo" else st.session_state.usuario_actual
    with st.form("f_c"):
        f_f = st.date_input("Fecha", value=datetime.now())
        f_t = st.selectbox("Tarea", list(COLORES_TAREAS.keys()))
        f_h = st.number_input("Horas", min_value=0.0, step=0.5)
        f_n = st.text_area("Nota")
        if st.form_submit_button("Guardar"):
            if f_h > 0:
                nueva = {'Fecha': f_f, 'Tarea': f_t, 'Nota': f_n}
                for op in OPERARIOS_FIJOS: nueva[op] = round(f_h, 2) if op == u_c else 0[cite: 1]
                st.session_state.cargas = pd.concat([st.session_state.cargas, pd.DataFrame([nueva])], ignore_index=True)
                guardar_df("Cargas", st.session_state.cargas)
                st.success("Guardado correctamente.")
                st.rerun()

# ===== 10. EXCEPCIONES Y RESET =====
elif menu == "Excepciones":
    st.title("Excepciones")
    with st.form("f_exc"):
        e_o = st.selectbox("Operario", OPERARIOS_FIJOS)
        e_i = st.date_input("Desde"); e_f = st.date_input("Hasta")
        e_m = st.selectbox("Motivo", ["Vacaciones", "Licencia", "Examen", "Otro"])
        if st.form_submit_button("Guardar"):
            hs_e = calcular_dias_habiles(e_i, e_f) * HORAS_DIA_LABORAL
            nueva_e = pd.DataFrame([{'Operario': e_o, 'Fecha Inicio': e_i, 'Fecha Fin': e_f, 'Motivo': e_m, 'Horas': hs_e}])
            st.session_state.excepciones = pd.concat([st.session_state.excepciones, nueva_e], ignore_index=True)
            guardar_df("Excepciones", st.session_state.excepciones)
            st.rerun()

elif menu == "Resetear Datos":
    st.title("⚠️ Resetear Datos")
    if st.text_input("Escriba 'BORRAR'") == "BORRAR":
        if st.button("Borrar todo"):
            df_v = pd.DataFrame(columns=['Fecha', 'Tarea'] + OPERARIOS_FIJOS + ['Nota'])
            guardar_df("Cargas", df_v)
            st.session_state.cargas = df_v
            st.rerun()
