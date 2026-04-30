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
st.set_page_config(page_title="CRM Capacidad Instalada", layout="wide")

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

TAREAS_DISPONIBLE_TIPO = [
    "DISPONIBLE",
    "PLANIFICACIONES/ORGANIZACIÓN/PROCEDIMIENTO S/INFORMES",
    "INASISTENCIA POR EXAMEN O TRAMITE"
]

MESES_ES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
    7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
}

DIAS_SEMANA_ES = {
    0: "Lunes", 1: "Martes", 2: "Miércoles", 3: "Jueves",
    4: "Viernes", 5: "Sábado", 6: "Domingo"
}

# ===== 2. FERIADOS (CSV) =====
@st.cache_data
def cargar_feriados():
    try:
        df_f = pd.read_csv("feriados_2026.csv")
        df_f.columns = df_f.columns.str.strip().str.lower()
        df_f['fecha'] = pd.to_datetime(df_f['fecha'], errors='coerce')
        return df_f['fecha'].dt.date.dropna().tolist()
    except: return []

FERIADOS = cargar_feriados()

# ===== 3. PDF =====
def generar_pdf(nombre, mes, anio, total_hs, eficiencia, estado, df_tareas):
    from reportlab.lib.pagesizes import letter
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib import colors
    buf = BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=letter)
    s = getSampleStyleSheet()
    story = [Paragraph(f"Reporte {nombre} - {MESES_ES[mes]} {anio}", s['Title']), Spacer(1,12)]
    data = [["Métrica", "Valor"], ["Total", f"{total_hs}hs"], ["Eficiencia", f"{eficiencia:.1f}%"], ["Estado", estado]]
    t = Table(data, colWidths=[200, 200])
    t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.grey),('GRID',(0,0),(-1,-1),1,colors.black)]))
    story.append(t)
    doc.build(story)
    buf.seek(0)
    return buf

# ===== 4. GOOGLE SHEETS =====
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
        ws.clear()
        df_c = df.copy()
        for c in df_c.columns:
            if 'fecha' in c.lower(): df_c[c] = pd.to_datetime(df_c[c], errors='coerce').dt.strftime('%d/%m/%Y')
        ws.update([df_c.columns.values.tolist()] + df_c.fillna('').astype(str).values.tolist())
        st.cache_data.clear()
        return True
    except: return False

# ===== 5. LÓGICA CAPACIDAD =====
def calcular_dias(ini, fin):
    return len(pd.bdate_range(start=ini, end=fin, freq='C', holidays=FERIADOS))

# ===== 6. INICIO APP =====
if 'cargas' not in st.session_state:
    df, _ = cargar_hoja("Cargas")
    st.session_state.cargas = df if not df.empty else pd.DataFrame(columns=['Fecha', 'Tarea'] + OPERARIOS_FIJOS + ['Nota'])

if 'usuario_actual' not in st.session_state:
    st.session_state.usuario_actual = None

if st.session_state.usuario_actual is None:
    st.title("CRM Estudio Contable")
    u = st.selectbox("Usuario:", ["Seleccionar..."] + OPERARIOS_FIJOS + ["Admin - Ver todo"])
    if st.button("Ingresar") and u != "Seleccionar...":
        st.session_state.usuario_actual = u
        st.rerun()
    st.stop()

# ===== 7. MENÚ SIDEBAR =====
opciones = ["Panel de Control", "Cargar Horas", "Carga Masiva", "Resetear Datos"] if st.session_state.usuario_actual == "Admin - Ver todo" else ["Panel de Control", "Cargar Mis Horas"]
menu = st.sidebar.radio("Menú", opciones)

if st.sidebar.button("Salir"):
    st.session_state.clear()
    st.rerun()

# ===== 8. PANEL DE CONTROL =====
if menu == "Panel de Control":
    st.title("Panel de Control - Ocupación")
    c1, c2 = st.columns(2)
    with c1: anio = st.selectbox("Año", [2025, 2026, 2027], index=1)
    with c2: mes = st.selectbox("Mes", list(range(1,13)), format_func=lambda x: MESES_ES[x], index=datetime.now().month-1)

    inicio_m = datetime(anio, mes, 1).date()
    fin_m = (datetime(anio, mes+1, 1) if mes < 12 else datetime(anio+1, 1, 1)).date() - timedelta(days=1)
    dias_h = calcular_dias(inicio_m, fin_m)
    cap_base = dias_h * HORAS_DIA_LABORAL
    st.info(f"**{MESES_ES[mes]} {anio}**: {dias_h} días hábiles. Capacidad: {cap_base} hs.")

    df_p = st.session_state.cargas.copy()
    if not df_p.empty:
        df_p['Fecha'] = pd.to_datetime(df_p['Fecha'], errors='coerce')
        df_m = df_p[(df_p['Fecha'].dt.month == mes) & (df_p['Fecha'].dt.year == anio)]
        
        # --- SECCIÓN INDIVIDUAL ---
        p_sel = st.selectbox("Operario:", OPERARIOS_FIJOS) if st.session_state.usuario_actual == "Admin - Ver todo" else st.session_state.usuario_actual
        df_ind_melt = df_m.melt(id_vars=['Fecha', 'Tarea'], value_vars=[p_sel], var_name='Op', value_name='Hs')
        df_ind_melt = df_ind_melt[df_ind_melt['Hs'] > 0]
        
        res_t = df_ind_melt.groupby('Tarea')['Hs'].sum().round(1).reset_index()
        total_c = res_t['Hs'].sum().round(1)
        hs_p = df_ind_melt[~df_ind_melt['Tarea'].isin(TAREAS_DISPONIBLE_TIPO)]['Hs'].sum()
        eficiencia = (hs_p / cap_base * 100) if cap_base > 0 else 0
        hs_l = df_ind_melt[df_ind_melt['Tarea'].isin(TAREAS_DISPONIBLE_TIPO)]['Hs'].sum()
        porc_l = (hs_l / total_c * 100) if total_c > 0 else 0
        
        if porc_l >= 20.0: est = "🟢 OK"
        elif porc_l >= 10.0: est = "🟡 Atención"
        else: est = "🔴 Al límite"

        col_g, col_m = st.columns([2,1])
        with col_g:
            fig = px.pie(res_t, values='Hs', names='Tarea', color='Tarea', color_discrete_map=COLORES_TAREAS, title=f"Ocupación {p_sel}")
            st.plotly_chart(fig, use_container_width=True)
        with col_m:
            st.metric("Total Cargado", f"{total_c} hs")
            st.metric("Eficiencia", f"{eficiencia:.1f}%")
            st.metric("Disponibilidad", f"{porc_l:.1f}%")
            st.subheader(f"Estado: {est}")
            if st.button("Generar Reporte PDF"):
                pdf = generar_pdf(p_sel, mes, anio, total_c, eficiencia, est, res_t)
                st.download_button("Descargar PDF", pdf, f"Reporte_{p_sel}.pdf")

        # --- HISTÓRICO 3 MESES ---
        st.divider()
        st.subheader("Histórico Últimos 3 Meses")
        h_data = []
        for i in range(3):
            m_h = mes - i
            a_h = anio
            if m_h <= 0: m_h += 12; a_h -= 1
            df_h = df_p[(df_p['Fecha'].dt.month == m_h) & (df_p['Fecha'].dt.year == a_h)]
            total_h = df_h[p_sel].sum()
            h_data.append({'Mes': f"{MESES_ES[m_h][:3]}", 'Hs': total_h})
        st.plotly_chart(px.bar(pd.DataFrame(h_data[::-1]), x='Mes', y='Hs', title="Evolución Horas"), use_container_width=True)

        # --- EQUIPO TOTAL (ADMIN) ---
        if st.session_state.usuario_actual == "Admin - Ver todo":
            st.divider()
            st.subheader("Visión General del Equipo")
            df_eq = df_m.melt(id_vars=['Fecha', 'Tarea'], value_vars=OPERARIOS_FIJOS, var_name='Op', value_name='Hs')
            res_eq = df_eq.groupby('Tarea')['Hs'].sum().round(1).reset_index()
            st.plotly_chart(px.pie(res_eq, values='Hs', names='Tarea', color='Tarea', color_discrete_map=COLORES_TAREAS, title="Distribución Total Equipo"), use_container_width=True)
            res_ops = []
            for o in OPERARIOS_FIJOS:
                res_ops.append({'Operario': o, 'Total Hs': df_m[o].sum().round(1)})
            st.table(pd.DataFrame(res_ops))

# ===== 9. CARGA MASIVA =====
elif menu == "Carga Masiva":
    st.title("Carga Masiva de Horas")
    with st.form("f_masiva"):
        u_m = st.selectbox("Operario", OPERARIOS_FIJOS)
        t_m = st.selectbox("Tarea", list(COLORES_TAREAS.keys()))
        f_i = st.date_input("Inicio")
        f_f = st.date_input("Fin")
        h_t = st.number_input("Horas Totales", min_value=0.0)
        if st.form_submit_button("Guardar"):
            dias = pd.bdate_range(start=f_i, end=f_f, freq='C', holidays=FERIADOS)
            if len(dias) > 0:
                h_d = round(h_t / len(dias), 2)
                filas = []
                for d in dias:
                    f = {'Fecha': d, 'Tarea': t_m, 'Nota': 'Carga Masiva'}
                    for o in OPERARIOS_FIJOS: f[o] = h_d if o == u_m else 0
                    filas.append(f)
                st.session_state.cargas = pd.concat([st.session_state.cargas, pd.DataFrame(filas)], ignore_index=True)
                guardar_df("Cargas", st.session_state.cargas)
                st.success("Carga masiva completada")
                st.rerun()

# ===== 10. CARGAR HORAS Y REGISTROS (HISTORIAL) =====
elif "Cargar" in menu:
    st.title("Cargar Horas")
    u_c = st.selectbox("Persona:", OPERARIOS_FIJOS) if st.session_state.usuario_actual == "Admin - Ver todo" else st.session_state.usuario_actual
    
    with st.form("f_ind"):
        f_f = st.date_input("Fecha", value=datetime.now())
        f_t = st.selectbox("Tarea", list(COLORES_TAREAS.keys()))
        f_h = st.number_input("Horas", step=0.5)
        f_n = st.text_input("Nota")
        if st.form_submit_button("Guardar"):
            nueva = {'Fecha': f_f, 'Tarea': f_t, 'Nota': f_n}
            for op in OPERARIOS_FIJOS: nueva[op] = f_h if op == u_c else 0
            st.session_state.cargas = pd.concat([st.session_state.cargas, pd.DataFrame([nueva])], ignore_index=True)
            guardar_df("Cargas", st.session_state.cargas)
            st.success("Guardado")
            st.rerun()

    st.divider()
    
    # --- CUADRO DE RESUMEN POR MES (CON ORDEN CRONOLÓGICO) ---
    st.subheader(f"Resumen Mensual por Tarea para {u_c}")
    df_res = st.session_state.cargas.copy()
    df_res['Fecha'] = pd.to_datetime(df_res['Fecha'], errors='coerce')
    df_res = df_res[df_res[u_c] > 0]
    
    if not df_res.empty:
        df_res['Mes_N'] = df_res['Fecha'].dt.month
        df_res['Año'] = df_res['Fecha'].dt.year
        df_res['Mes'] = df_res['Mes_N'].map(MESES_ES)
        
        # Agrupamos incluyendo el número del mes para poder ordenar correctamente
        cuadro = df_res.groupby(['Año', 'Mes_N', 'Mes', 'Tarea'])[u_c].sum().reset_index()
        cuadro.columns = ['Año', 'Mes_N', 'Mes', 'Tarea', 'Hs Totales']
        
        # Ordenamos por Año y Número de Mes de forma descendente (Abril -> Marzo -> Febrero)
        cuadro_ordenado = cuadro.sort_values(by=['Año', 'Mes_N'], ascending=False)
        
        # Quitamos la columna auxiliar Mes_N antes de mostrar la tabla
        st.dataframe(cuadro_ordenado[['Año', 'Mes', 'Tarea', 'Hs Totales']], use_container_width=True, hide_index=True)

    st.divider()
    
    # --- HISTORIAL DETALLADO ---
    st.subheader(f"Historial de Cargas Detallado para {u_c}")
    ver_solo_dia = st.checkbox("Ver solo el día seleccionado arriba")
    df_hist = st.session_state.cargas.copy()
    df_hist['Fecha'] = pd.to_datetime(df_hist['Fecha'], errors='coerce')
    df_hist = df_hist[df_hist[u_c] > 0]
    
    if ver_solo_dia:
        df_hist = df_hist[df_hist['Fecha'].dt.date == f_f]
    
    df_hist = df_hist.sort_values('Fecha', ascending=False)
    
    if df_hist.empty:
        st.info("No hay registros cargados.")
    else:
        for i, row in df_hist.iterrows():
            with st.container():
                col_info, col_del = st.columns([5, 1])
                fecha_str = row['Fecha'].strftime('%d/%m/%Y')
                dia_sem = DIAS_SEMANA_ES[row['Fecha'].weekday()]
                col_info.write(f"**{fecha_str} ({dia_sem})** - {row['Tarea']}: **{row[u_c]} hs** | *{row.get('Nota', '')}*")
                if col_del.button("Eliminar", key=f"del_{i}"):
                    st.session_state.cargas = st.session_state.cargas.drop(i).reset_index(drop=True)
                    guardar_df("Cargas", st.session_state.cargas)
                    st.rerun()

elif menu == "Resetear Datos":
    if st.text_input("Escriba BORRAR") == "BORRAR":
        if st.button("Eliminar Todo"):
            guardar_df("Cargas", pd.DataFrame(columns=['Fecha', 'Tarea'] + OPERARIOS_FIJOS + ['Nota']))
            st.rerun()
