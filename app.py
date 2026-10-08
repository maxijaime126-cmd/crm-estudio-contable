"""CRM Grupo Pressacco v2 — capacidad instalada por departamento.

Departamento -> Tarea -> Subtarea, minutos, ausencias, horas extra automáticas, calendario de
saturación, vista semanal por departamento, desvío y tendencia, y PDFs.
La lógica está en calc.py (con pruebas), los datos en store.py y los PDFs en pdf.py.
"""
import hmac
import html
from datetime import date, datetime
from zoneinfo import ZoneInfo

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

import calc
import config as C
import pdf
import store as S

st.set_page_config(page_title="Grupo Pressacco · Capacidad", layout="wide", page_icon="🏛️")

st.markdown("""
<style>
    :root { --azul:#1F4E78; --celeste:#0077B6; --turq:#00B4D8; --borde:#D6E4F0; }
    .block-container { padding-top: 1.4rem; max-width: 1250px; }
    h1, h2, h3 { color: var(--azul); }
    /* Barra lateral */
    section[data-testid="stSidebar"] { background: linear-gradient(180deg, #DCEBF7 0%, #F2F8FD 100%);
                                       border-right: 1px solid var(--borde); }
    /* Encabezado */
    .hero { background: linear-gradient(120deg, #1F4E78 0%, #0077B6 60%, #00B4D8 100%); border-radius: 16px;
            padding: 18px 24px; margin-bottom: 18px; display: flex; justify-content: space-between;
            align-items: center; flex-wrap: wrap; gap: 10px; box-shadow: 0 4px 14px rgba(31,78,120,.25); }
    .hero-t { color: #fff; font-size: 1.55rem; font-weight: 800; letter-spacing: .3px; }
    .hero-s { color: #D8EEF9; font-size: .95rem; }
    .hero-chip { background: rgba(255,255,255,.18); color: #fff; padding: 6px 14px; border-radius: 999px;
                 font-size: .9rem; font-weight: 600; }
    /* Tarjetas de indicadores */
    .kpi-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap: 12px;
                margin: 4px 0 16px 0; }
    .kpi-card { border-radius: 14px; padding: 14px 12px; color: #fff; text-align: center;
                box-shadow: 0 3px 10px rgba(0,0,0,.12); }
    .kpi-v { font-size: 1.75rem; font-weight: 800; color: #fff; white-space: nowrap; line-height: 1.25; text-align: center; }
    .kpi-card p { font-size: .8rem; margin: 3px 0 0 0; opacity: .95; }
    .k-cap { background: linear-gradient(135deg, #1F4E78, #3A7CA5); }
    .k-trab { background: linear-gradient(135deg, #2D9C6B, #52C48F); }
    .k-util { background: linear-gradient(135deg, #6C4AB6, #9B7FDB); }
    .k-disp { background: linear-gradient(135deg, #0096C7, #48CAE4); }
    .k-aus { background: linear-gradient(135deg, #6B7C8C, #98A8B8); }
    .k-extra { background: linear-gradient(135deg, #E76F51, #F4A261); }
    /* Alertas y tarjetas */
    .alerta-box { border-left: 5px solid #E76F51; background: #FFF4EE; border-radius: 8px;
                  padding: 10px 15px; margin-bottom: 8px; font-size: .92rem; color: #333; }
    .obj-box { background: #fff; border: 1px solid var(--borde); border-radius: 12px; padding: 12px 16px;
               margin-bottom: 6px; color: #1B2A38; }
    .dep-card { border-radius: 12px; padding: 12px 16px; margin: 6px 0 14px 0; font-size: 1rem; color: #1B2A38; }
    .chip { display: inline-block; padding: 2px 10px; border-radius: 999px; font-size: .78rem; font-weight: 700;
            margin-left: 8px; color: #fff; }
    .chip-trab { background: #2D9C6B; } .chip-disp { background: #0096C7; } .chip-aus { background: #6B7C8C; }
    .sec { font-size: 1.15rem; font-weight: 700; color: var(--azul); margin: 18px 0 4px 0; }
    div.stButton > button[kind="primary"] { border-radius: 10px; font-weight: 700; padding: .5rem 1.4rem; }
</style>
""", unsafe_allow_html=True)

DIAS_ES = ["lun", "mar", "mié", "jue", "vie", "sáb", "dom"]
CHIP_TIPO = {C.TIPO_TRABAJO: "chip-trab", C.TIPO_DISPONIBLE: "chip-disp", C.TIPO_AUSENCIA: "chip-aus"}
COLOR_MAS, COLOR_MENOS = "#F4A261", "#4C9AD6"


# ============================================================================
# Utilidades
# ============================================================================
def hoy_ar() -> date:
    return datetime.now(ZoneInfo(C.TZ)).date()


def fmt_hs(x: float) -> str:
    return f"{x:.1f}".replace(".", ",")


def fmt_min(m: int) -> str:
    h, mm = divmod(int(m), 60)
    return f"{h} h {mm:02d} min" if h else f"{mm} min"


def etiqueta_depto(d: str) -> str:
    return f"{C.ICONOS_DEPTO.get(d, '📁')} {d}"


def seccion(texto: str):
    st.markdown(f'<div class="sec">{html.escape(texto)}</div>', unsafe_allow_html=True)


def hero(usuario: str, es_admin: bool, hoy: date):
    rol = "Administrador" if es_admin else "Equipo"
    st.markdown(
        f'<div class="hero"><div><div class="hero-t">🏛️ Grupo Pressacco</div>'
        f'<div class="hero-s">Capacidad instalada · {C.MESES_ES[hoy.month]} {hoy.year}</div></div>'
        f'<div class="hero-chip">👤 {html.escape(usuario)} · {rol}</div></div>', unsafe_allow_html=True)


def kpis(items: list[tuple]):
    """items: (valor, etiqueta, clase). Una sola pieza de HTML que se adapta al ancho."""
    cards = "".join(
        f'<div class="kpi-card {html.escape(c)}"><div class="kpi-v">{html.escape(str(v))}</div><p>{html.escape(l)}</p></div>'
        for v, l, c in items)
    st.markdown(f'<div class="kpi-grid">{cards}</div>', unsafe_allow_html=True)


def alertas(mensajes: list[str]):
    """Un solo bloque (no un elemento por alerta): evita errores de renderizado."""
    if mensajes:
        cuerpo = "".join(f'<div class="alerta-box">{html.escape(m)}</div>' for m in mensajes)
        st.markdown(cuerpo, unsafe_allow_html=True)


def estilo_fig(fig, alto=None):
    fig.update_layout(plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                      margin=dict(l=0, r=0, t=30, b=0), font=dict(family="sans-serif", size=12))
    if alto:
        fig.update_layout(height=alto)
    return fig


def boton_pdf(clave: str, etiqueta: str, generar, nombre: str):
    """Genera el PDF al tocar el botón (no en cada recarga) y ofrece la descarga."""
    if st.button(etiqueta, key=f"btn_{clave}"):
        try:
            st.session_state[f"pdf_{clave}"] = generar()
        except Exception as e:  # noqa: BLE001
            st.error(f"No se pudo generar el PDF: {e}")
    datos = st.session_state.get(f"pdf_{clave}")
    if datos:
        st.download_button("⬇️ Descargar PDF", data=datos, file_name=nombre, mime="application/pdf",
                           key=f"dl_{clave}")


# ============================================================================
# Datos (con caché corto: lo que carga otra persona se ve en segundos)
# ============================================================================
@st.cache_resource(show_spinner="Conectando con Google Sheets…")
def get_store():
    import gspread
    from google.oauth2.service_account import Credentials
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=["https://www.googleapis.com/auth/spreadsheets",
                "https://www.googleapis.com/auth/drive"])
    store = S.SheetsStore(gspread.authorize(creds).open(C.SPREADSHEET))
    store.bootstrap()   # crea hojas y columnas que falten
    return store


@st.cache_data(ttl=20, show_spinner=False)
def leer(hoja: str) -> pd.DataFrame:
    return get_store().leer(hoja)


def invalidar():
    st.cache_data.clear()


def cargar_contexto() -> dict:
    cat_raw = leer(C.HOJA_CATALOGO)
    return {
        "cat": calc.catalogo_activo(cat_raw),
        "personas": calc.personas_df(leer(C.HOJA_PERSONAS)),
        "feriados": calc.feriados_set(leer(C.HOJA_FERIADOS)),
        "reg": calc.preparar_registros(leer(C.HOJA_REGISTROS), cat_raw),
    }


# ============================================================================
# Login
# ============================================================================
def pantalla_login(personas: pd.DataFrame):
    st.markdown('<div class="hero"><div><div class="hero-t">🏛️ Grupo Pressacco</div>'
                '<div class="hero-s">Capacidad instalada del estudio</div></div></div>', unsafe_allow_html=True)
    _, centro, _ = st.columns([1, 2, 1])
    with centro:
        activos = personas[personas["Activo"]]["Nombre"].tolist()
        u = st.selectbox("Usuario:", ["Seleccionar..."] + activos)
        pwd = st.text_input("Contraseña:", type="password")
        if st.button("Ingresar", type="primary") and u != "Seleccionar...":
            correcta = st.secrets.get("passwords", {}).get(u)
            if correcta is None:
                st.error("No hay contraseña configurada para este usuario. Pedile al Admin que la agregue en Secrets.")
            elif hmac.compare_digest(str(pwd).encode(), str(correcta).encode()):
                st.session_state["usuario"] = u
                st.rerun()
            else:
                st.error("Contraseña incorrecta.")


# ============================================================================
# Cargar horas
# ============================================================================
def banner_objetivo(reg, persona, hd, feriados, hoy):
    objetivo, cargadas = calc.objetivo_mes(reg, persona, hoy.year, hoy.month, hd, feriados)
    faltan = max(0.0, objetivo - cargadas)
    st.markdown(
        f'<div class="obj-box">🎯 <b>Objetivo {C.MESES_ES[hoy.month]}:</b> faltan <b>{fmt_hs(faltan)} hs</b> '
        f'para las {fmt_hs(objetivo)} hs del mes · cargadas {fmt_hs(cargadas)} hs</div>', unsafe_allow_html=True)
    st.progress(min(1.0, cargadas / objetivo) if objetivo > 0 else 0.0)


def pantalla_carga(ctx: dict, usuario: str, es_admin: bool):
    cat, reg, feriados, personas = ctx["cat"], ctx["reg"], ctx["feriados"], ctx["personas"]
    horas_pp = calc.horas_por_persona(personas)
    hoy = hoy_ar()
    st.header("➕ Cargar horas")
    if "msg_ok" in st.session_state:
        st.success(st.session_state.pop("msg_ok"))
    if cat.empty:
        st.error("El catálogo está vacío. Revisá la hoja Catalogo del Sheet.")
        return

    persona = st.selectbox("Persona:", list(horas_pp)) if es_admin and horas_pp else usuario
    hd = horas_pp.get(persona, C.HORAS_DIA_DEFAULT)
    banner_objetivo(reg, persona, hd, feriados, hoy)

    # --- Qué se carga ---
    modo_carga = st.radio("¿Qué vas a cargar?", ["💼 Trabajo", "🟦 Disponible", "🏖️ Ausencia"],
                          horizontal=True, key="c_modo")
    if modo_carga.startswith("💼"):
        # Disponible y Ausencias tienen su propia opción: acá van las tareas de cada departamento
        cat_t = cat[~cat["Tarea"].isin([C.TAREA_DISPONIBLE, C.TAREA_AUSENCIAS])]
        if cat_t.empty:
            st.error("El catálogo no tiene tareas de trabajo. Revisá la hoja Catalogo.")
            return
        deptos = list(dict.fromkeys(cat_t["Departamento"]))
        depto = st.selectbox("Departamento", deptos, format_func=etiqueta_depto, key="c_dep")
        sub_d = cat_t[cat_t["Departamento"] == depto]
        c2, c3 = st.columns(2)
        tarea = c2.selectbox("Tarea", list(dict.fromkeys(sub_d["Tarea"])), key=f"c_tar_{depto}")
        subs = [x for x in sub_d[sub_d["Tarea"] == tarea]["Subtarea"] if x]
        if subs:
            subtarea = c3.selectbox("Subtarea", subs, key=f"c_sub_{depto}_{tarea}")
        else:
            subtarea = ""
            c3.caption("Esta tarea no tiene subtareas.")
        fila_cat = sub_d[(sub_d["Tarea"] == tarea) & (sub_d["Subtarea"] == subtarea)]
        tipo = fila_cat["Tipo"].iloc[0] if not fila_cat.empty else C.TIPO_TRABAJO
        color, icono = C.COLORES_DEPTO.get(depto, "#0077B6"), C.ICONOS_DEPTO.get(depto, "📁")
        ruta = f"<b>{html.escape(depto)}</b> › {html.escape(tarea)}" + (f" › {html.escape(subtarea)}" if subtarea else "")
    elif modo_carga.startswith("🟦"):
        depto, tarea, subtarea, tipo = C.DEPTO_GENERAL, C.TAREA_DISPONIBLE, "", C.TIPO_DISPONIBLE
        color, icono, ruta = "#0096C7", "🟦", "<b>Disponible</b>"
        st.caption("Para cuando no estás haciendo nada: completa tus horas del día. No hace falta elegir departamento.")
    else:
        opciones = list(dict.fromkeys(x for x in cat[cat["Tarea"] == C.TAREA_AUSENCIAS]["Subtarea"] if x)) \
            or ["Inasistencia (día personal)", "Vacaciones"]
        depto, tarea, tipo = C.DEPTO_GENERAL, C.TAREA_AUSENCIAS, C.TIPO_AUSENCIA
        subtarea = st.selectbox("Tipo de ausencia", opciones, key="c_aus")
        color, icono, ruta = "#6B7C8C", "🏖️", f"<b>Ausencia</b> › {html.escape(subtarea)}"
    es_ausencia = tipo == C.TIPO_AUSENCIA

    st.markdown(
        f'<div class="dep-card" style="border-left:8px solid {color}; background:{color}22;">'
        f'{icono} {ruta}<span class="chip {CHIP_TIPO.get(tipo, "chip-trab")}">'
        f'{html.escape(tipo)}</span></div>', unsafe_allow_html=True)

    # --- Cuándo y cuánto ---
    if es_ausencia:
        st.caption("Las ausencias bajan tu capacidad del mes; no cuentan como trabajo.")
        d1, d2 = st.columns(2)
        desde = d1.date_input("Desde", value=hoy, format="DD/MM/YYYY", key="a_desde")
        hasta = d2.date_input("Hasta", value=hoy, format="DD/MM/YYYY", key="a_hasta")
        completo = st.checkbox(f"Día completo ({fmt_hs(hd)} hs por día)", value=True, key="a_comp")
        if completo:
            minutos = int(hd * 60)
        else:
            minutos = st.number_input("Minutos por día", min_value=C.MIN_MINUTOS,
                                      max_value=int(hd * 60), step=C.PASO_MINUTOS, value=120, key="a_min")
        modo, solo_habiles = "por_dia", True
    else:
        modalidad = st.radio("¿Cuándo?", ["Un día", "Repartir en varios días"], horizontal=True)
        if modalidad == "Un día":
            desde = hasta = st.date_input("Fecha", value=hoy, max_value=hoy, format="DD/MM/YYYY", key="t_fecha")
            minutos = st.number_input("Minutos", min_value=C.MIN_MINUTOS, max_value=C.MAX_MINUTOS_CARGA,
                                      step=C.PASO_MINUTOS, value=60, key="t_min")
            modo = "por_dia"
        else:
            d1, d2 = st.columns(2)
            desde = d1.date_input("Desde", value=hoy, max_value=hoy, format="DD/MM/YYYY", key="t_desde")
            hasta = d2.date_input("Hasta", value=hoy, max_value=hoy, format="DD/MM/YYYY", key="t_hasta")
            minutos = st.number_input("Minutos totales a repartir", min_value=C.MIN_MINUTOS,
                                      max_value=60 * 200, step=C.PASO_MINUTOS, value=600, key="t_tot")
            modo = "total"
        solo_habiles = False
        st.caption(f"= {fmt_min(minutos)}")

    obligatoria = subtarea == C.SUBTAREA_NOTA_OBLIGATORIA
    nota = st.text_input("Nota (obligatoria: contá qué fue)" if obligatoria else "Nota (opcional)", key="c_nota")

    # --- Vista previa ---
    filas = []
    if hasta < desde:
        st.error("La fecha 'Hasta' no puede ser anterior a 'Desde'.")
    else:
        filas = calc.generar_registros(persona, depto, tarea, subtarea, nota.strip(), desde, hasta,
                                       feriados, minutos=int(minutos), modo=modo, solo_habiles=solo_habiles)
        if not filas:
            st.warning("En ese rango no hay días hábiles (fines de semana o feriados).")
        elif len(filas) == 1:
            st.info(f"Se va a cargar {fmt_min(filas[0]['Minutos'])} el {desde:%d/%m/%Y}.")
        else:
            total = sum(f["Minutos"] for f in filas) / 60
            st.info(f"Se van a cargar {len(filas)} días ({desde:%d/%m} al {hasta:%d/%m}): {fmt_hs(total)} hs en total.")
        if len(filas) == 1 and tipo == C.TIPO_TRABAJO:
            rp = reg[(reg["Persona"] == persona) & (reg["Fecha"].dt.date == desde)]
            d = calc.por_dia(rp, hd, feriados, desde, desde).iloc[0]
            trabajo = float(d[C.TIPO_TRABAJO]) + filas[0]["Minutos"] / 60
            extra = max(0.0, trabajo - float(d["Capacidad"]))
            st.caption(f"Ese día ya tenías {fmt_hs(float(d[C.TIPO_TRABAJO]))} hs de trabajo cargadas.")
            if extra > 0:
                st.warning(f"Con esta carga ese día quedan {fmt_hs(extra)} hs extra.")

    if st.button("💾 Guardar", type="primary", disabled=not filas):
        if obligatoria and not nota.strip():
            st.error("Para «Otros imprevistos» la nota es obligatoria.")
        else:
            guardado, error = False, None
            try:
                ahora = S.ahora_iso()
                for f in filas:
                    f.update({"ID": S.nuevo_id(), "Registrado": ahora})
                get_store().agregar(C.HOJA_REGISTROS, filas)
                invalidar()
                guardado = True
            except Exception as e:  # noqa: BLE001
                error = e
            if guardado:
                st.session_state["msg_ok"] = f"✅ Guardado: {len(filas)} registro(s) de {persona}."
                st.rerun()
            else:
                st.error(f"No se pudo guardar: {error}")

    seccion_mis_cargas(reg, persona, hoy)


def seccion_mis_cargas(reg: pd.DataFrame, persona: str, hoy: date):
    st.divider()
    seccion("🗂️ Cargas del mes")
    propias = reg[reg["Persona"] == persona]
    anios = sorted({hoy.year, *propias["Fecha"].dt.year.astype(int).tolist()})
    f1, f2, _ = st.columns([1.3, 1.6, 2.5])
    anio = f1.selectbox("Año", anios, index=anios.index(hoy.year), key="mc_anio")
    mes = f2.selectbox("Mes", list(range(1, 13)), index=hoy.month - 1,
                       format_func=lambda m: C.MESES_ES[m], key="mc_mes")
    mis = calc.del_mes(propias, anio, mes).sort_values(["Fecha", "Registrado"], ascending=False)
    if mis.empty:
        st.caption(f"No hay cargas de {persona} en {C.MESES_ES[mes]} {anio}.")
        return
    st.caption(f"{len(mis)} carga(s) · {fmt_hs(float(mis['Horas'].sum()))} hs en {C.MESES_ES[mes]} {anio}")
    st.dataframe(pd.DataFrame({
        "Fecha": mis["Fecha"].dt.strftime("%d/%m/%Y"), "Departamento": mis["Departamento"],
        "Tarea": mis["Tarea"], "Subtarea": mis["Subtarea"],
        "Minutos": mis["Minutos"].astype(int), "Nota": mis["Nota"]}), hide_index=True)

    editables = mis[mis["ID"] != ""]
    if editables.empty:
        return
    with st.expander("✏️ Editar o eliminar una carga"):
        etiquetas = {r.ID: (f"{r.Fecha:%d/%m} · {r.Departamento} › {r.Tarea}"
                            f"{' › ' + r.Subtarea if r.Subtarea else ''} · {int(r.Minutos)} min")
                     for r in editables.itertuples()}
        elegido = st.selectbox("Carga", list(etiquetas), format_func=lambda i: etiquetas[i])
        fila = editables[editables["ID"] == elegido].iloc[0]
        valor = max(C.PASO_MINUTOS, min(int(fila["Minutos"]), C.MAX_MINUTOS_CARGA))
        nuevos = st.number_input("Minutos", min_value=C.PASO_MINUTOS, max_value=C.MAX_MINUTOS_CARGA,
                                 step=C.PASO_MINUTOS, value=valor, key=f"e_min_{elegido}")
        nueva_nota = st.text_input("Nota", value=fila["Nota"], key=f"e_nota_{elegido}")
        b1, b2 = st.columns(2)
        if b1.button("Guardar cambios", key=f"e_ok_{elegido}"):
            try:
                get_store().actualizar_por_id(C.HOJA_REGISTROS, elegido,
                                              {"Minutos": int(nuevos), "Nota": nueva_nota.strip()})
                invalidar()
                st.session_state["msg_ok"] = "✅ Carga actualizada."
            except Exception as e:  # noqa: BLE001
                st.error(f"No se pudo actualizar: {e}")
            else:
                st.rerun()
        confirmar = b2.checkbox("Confirmo que quiero eliminarla", key=f"e_conf_{elegido}")
        if b2.button("🗑️ Eliminar", disabled=not confirmar, key=f"e_del_{elegido}"):
            try:
                get_store().borrar_por_id(C.HOJA_REGISTROS, [elegido])
                invalidar()
                st.session_state["msg_ok"] = "🗑️ Carga eliminada."
            except Exception as e:  # noqa: BLE001
                st.error(f"No se pudo eliminar: {e}")
            else:
                st.rerun()


# ============================================================================
# Panel de control
# ============================================================================
def pantalla_resumen(ctx: dict, usuario: str, es_admin: bool):
    reg = ctx["reg"]
    hoy = hoy_ar()
    st.header("📊 Panel de control")
    anios = sorted({hoy.year - 1, hoy.year, hoy.year + 1, *reg["Fecha"].dt.year.astype(int).tolist()})
    f1, f2, _ = st.columns([1.3, 1.6, 2.5])
    anio = f1.selectbox("Año", anios, index=anios.index(hoy.year))
    mes = f2.selectbox("Mes", list(range(1, 13)), index=hoy.month - 1, format_func=lambda m: C.MESES_ES[m])

    if es_admin:
        t1, t2, t3, t4 = st.tabs(["👤 Individual", "🌐 Equipo", "🗓️ Calendario", "📆 Semanal"])
        with t1:
            tab_individual(ctx, None, anio, mes, hoy, es_admin)
        with t2:
            tab_equipo(ctx, anio, mes)
        with t3:
            tab_calendario(ctx, anio, mes, hoy)
        with t4:
            tab_semanal(ctx, anio, mes)
    else:
        tab_individual(ctx, usuario, anio, mes, hoy, es_admin)


def tab_individual(ctx, persona, anio, mes, hoy, es_admin):
    reg, feriados = ctx["reg"], ctx["feriados"]
    horas_pp = calc.horas_por_persona(ctx["personas"])
    if persona is None:
        if not horas_pp:
            st.warning("No hay operarios activos en la hoja Personas.")
            return
        persona = st.selectbox("Persona:", list(horas_pp), key="ind_persona")
    hd = horas_pp.get(persona, C.HORAS_DIA_DEFAULT)
    r = calc.resumen_persona_mes(reg, persona, anio, mes, hd, feriados)
    periodo = f"{C.MESES_ES[mes]} {anio}"

    kpis([
        (fmt_hs(r["capacidad"]), "Capacidad del mes (hs)", "k-cap"),
        (fmt_hs(r["trabajo"]), "Trabajo (hs)", "k-trab"),
        (f"{r['utilizacion']:.0f}%", "Utilización", "k-util"),
        (fmt_hs(r["disponible"]), "Disponible (hs)", "k-disp"),
        (f"{r['disponibilidad']:.0f}%", "Disponibilidad", "k-disp"),
        (fmt_hs(r["ausencia"]), "Ausencias (hs)", "k-aus"),
        (fmt_hs(r["extra"]), "Horas extra", "k-extra"),
    ])

    msgs = []
    falta = calc.dias_sin_carga(reg, persona, anio, mes, feriados, hoy)
    if falta:
        ej = ", ".join(f"{d:%d/%m}" for d in falta[:6]) + (" y más" if len(falta) > 6 else "")
        msgs.append(f"📅 {len(falta)} día(s) hábil(es) sin carga: {ej}.")
    if r["extra"] > 0:
        msgs.append(f"⏱️ {fmt_hs(r['extra'])} hs extra en el mes.")
    if es_admin:
        sin = calc.anios_sin_feriados(feriados, [anio])
        if sin:
            msgs.append(f"⚠️ No hay feriados cargados para {sin[0]}: la capacidad está sobreestimada. "
                        "Agregalos en la hoja Feriados.")
    alertas(msgs)

    ini, fin = calc.rango_mes(anio, mes)
    rm = reg[(reg["Persona"] == persona) & (reg["Fecha"].dt.date >= ini) & (reg["Fecha"].dt.date <= fin)]
    trabajo = rm[rm["Tipo"] == C.TIPO_TRABAJO]
    por_dep = pd.DataFrame(columns=["Departamento", "Horas"])
    detalle_pdf = pd.DataFrame(columns=["Departamento", "Tarea", "Subtarea", "Horas"])

    seccion("Horas de trabajo por departamento")
    if trabajo.empty:
        st.info("Sin horas de trabajo cargadas en este mes.")
    else:
        por_dep = trabajo.groupby("Departamento", as_index=False)["Horas"].sum().sort_values("Horas")
        fig = px.bar(por_dep, x="Horas", y="Departamento", orientation="h", color="Departamento",
                     color_discrete_map=C.COLORES_DEPTO, text=por_dep["Horas"].round(1))
        fig.update_layout(showlegend=False)
        st.plotly_chart(estilo_fig(fig))
        detalle_pdf = (trabajo.groupby(["Departamento", "Tarea", "Subtarea"], as_index=False)["Horas"].sum()
                       .sort_values(["Departamento", "Horas"], ascending=[True, False]))
        dep_sel = st.selectbox("Ver detalle de:", list(por_dep["Departamento"][::-1]),
                               format_func=etiqueta_depto, key="ind_dep")
        det = (detalle_pdf[detalle_pdf["Departamento"] == dep_sel][["Tarea", "Subtarea", "Horas"]]
               .sort_values("Horas", ascending=False).copy())
        det["% del departamento"] = (det["Horas"] / det["Horas"].sum() * 100).round(0)
        det["Horas"] = det["Horas"].round(1)
        st.dataframe(det, hide_index=True)

    # --- Desvío contra el promedio histórico ---
    seccion("Desvío contra el promedio de los 2 meses anteriores")
    nivel = st.radio("Nivel", ["Departamento", "Tarea"], horizontal=True, key="desvio_nivel")
    dv, usados = calc.desvio_historico(reg, persona, anio, mes, nivel)
    if dv.empty:
        st.caption("Todavía no hay horas de trabajo para comparar.")
    elif usados == 0:
        st.caption("No hay meses anteriores con datos todavía: el desvío aparece cuando haya historia.")
    else:
        st.caption(f"Promedio de {usados} mes(es) anterior(es) con datos. Naranja: más horas que el promedio. "
                   "Azul: menos horas. No es bueno ni malo: marca dónde cambió la carga.")
        dv = dv.copy()
        dv["Item"] = dv["Departamento"] if nivel == "Departamento" else dv["Departamento"] + " › " + dv["Tarea"]
        dv["Sentido"] = ["Más que el promedio" if x >= 0 else "Menos que el promedio" for x in dv["Desvío (hs)"]]
        top = dv.head(12).sort_values("Desvío (hs)")
        fig = px.bar(top, x="Desvío (hs)", y="Item", orientation="h", color="Sentido",
                     color_discrete_map={"Más que el promedio": COLOR_MAS, "Menos que el promedio": COLOR_MENOS})
        fig.update_layout(showlegend=False, yaxis_title=None)
        st.plotly_chart(estilo_fig(fig, max(260, 34 * len(top) + 60)))
        cols = ["Item", "Actual (hs)", "Promedio previo (hs)", "Desvío (hs)", "Desvío %"]
        st.dataframe(dv[cols], hide_index=True)

    # --- Tendencia y semanas ---
    seccion("Tendencia de los últimos 6 meses")
    tend, meses = calc.tendencia_departamentos(reg, persona, anio, mes, 6)
    if tend.empty:
        st.caption("Sin datos de trabajo en los últimos 6 meses.")
    else:
        fig = px.bar(tend, x="Mes", y="Horas", color="Departamento", color_discrete_map=C.COLORES_DEPTO,
                     category_orders={"Mes": meses})
        st.plotly_chart(estilo_fig(fig))

    seccion(f"Distribución semanal — {periodo}")
    sem, semanas = calc.distribucion_semanal(reg, persona, anio, mes)
    if sem.empty:
        st.caption("Sin horas de trabajo en este mes.")
    else:
        fig = px.bar(sem, x="Semana", y="Horas", color="Departamento", color_discrete_map=C.COLORES_DEPTO,
                     category_orders={"Semana": semanas})
        st.plotly_chart(estilo_fig(fig))

    libres = rm[rm["Tipo"] != C.TIPO_TRABAJO]
    if not libres.empty:
        with st.expander("Disponible y ausencias del mes"):
            tl = libres.groupby(["Tipo", "Departamento", "Tarea", "Subtarea"], as_index=False)["Horas"].sum()
            tl["Horas"] = tl["Horas"].round(1)
            st.dataframe(tl, hide_index=True)

    extras_pdf = None
    if r["extra"] > 0:
        d = r["detalle_dias"]
        d = d[d["Extra"] > 0].reset_index(names="Fecha")
        extras_pdf = pd.DataFrame({
            "Fecha": d["Fecha"].dt.strftime("%d/%m/%Y"), "Trabajo (hs)": d[C.TIPO_TRABAJO].round(1),
            "Capacidad del día (hs)": d["Capacidad"].round(1), "Extra (hs)": d["Extra"].round(1)})
        with st.expander("⏱️ Días con horas extra"):
            st.dataframe(extras_pdf, hide_index=True)

    seccion("Informe")
    boton_pdf(f"ind_{persona}_{anio}_{mes}", "📄 Preparar PDF mensual",
              lambda: pdf.pdf_individual(persona, periodo, r, por_dep, detalle_pdf, extras_pdf),
              f"Informe_{persona}_{anio}-{mes:02d}.pdf")


def datos_equipo(ctx, anio, mes):
    reg, feriados = ctx["reg"], ctx["feriados"]
    horas_pp = calc.horas_por_persona(ctx["personas"])
    filas = []
    for p, hd in horas_pp.items():
        r = calc.resumen_persona_mes(reg, p, anio, mes, hd, feriados)
        filas.append({"Persona": p, "Capacidad (hs)": r["capacidad"], "Trabajo (hs)": r["trabajo"],
                      "Utilización %": r["utilizacion"], "Disponible (hs)": r["disponible"],
                      "Disponibilidad %": r["disponibilidad"], "Ausencias (hs)": r["ausencia"],
                      "Extra (hs)": r["extra"]})
    return pd.DataFrame(filas)


def tab_equipo(ctx, anio, mes):
    tabla = datos_equipo(ctx, anio, mes)
    if tabla.empty:
        st.info("No hay operarios activos.")
        return
    seccion("Capacidad y carga por persona")
    st.dataframe(tabla, hide_index=True)
    kpis([(fmt_hs(tabla["Capacidad (hs)"].sum()), "Capacidad del equipo (hs)", "k-cap"),
          (fmt_hs(tabla["Trabajo (hs)"].sum()), "Trabajo del equipo (hs)", "k-trab"),
          (fmt_hs(tabla["Disponible (hs)"].sum()), "Disponible (hs)", "k-disp"),
          (fmt_hs(tabla["Extra (hs)"].sum()), "Horas extra (hs)", "k-extra")])

    reg, feriados, cat = ctx["reg"], ctx["feriados"], ctx["cat"]
    horas_pp = calc.horas_por_persona(ctx["personas"])
    deptos = list(dict.fromkeys(cat["Departamento"]))
    _, resumen_sem = calc.semanal_departamentos(reg, anio, mes, horas_pp, feriados, deptos)
    horas_d, cap_d = calc.matriz_departamentos(reg, anio, mes, horas_pp, feriados, deptos)
    picos = calc.top_dias_pico(horas_d, cap_d, 3)
    if not picos.empty:
        picos["Día"] = picos["Día"].map(lambda d: f"{DIAS_ES[d.weekday()]} {d:%d/%m}")
    seccion("Informe")
    boton_pdf(f"eq_{anio}_{mes}", "📄 Preparar PDF del equipo",
              lambda: pdf.pdf_equipo(f"{C.MESES_ES[mes]} {anio}", tabla, resumen_sem, picos),
              f"Informe_equipo_{anio}-{mes:02d}.pdf")


def tab_calendario(ctx, anio, mes, hoy):
    reg, feriados, cat = ctx["reg"], ctx["feriados"], ctx["cat"]
    horas_pp = calc.horas_por_persona(ctx["personas"])
    deptos = list(dict.fromkeys(cat["Departamento"]))
    horas, cap = calc.matriz_departamentos(reg, anio, mes, horas_pp, feriados, deptos)
    if horas.to_numpy().sum() == 0:
        st.info("Todavía no hay horas de trabajo cargadas en este mes.")
        return
    st.caption("Cada fila es un departamento. El color compara los días de ese mismo departamento "
               "(rojo = su día de mayor carga). Los números son las horas de trabajo cargadas.")
    dias = [d for d in horas.columns if calc.es_habil(d, feriados) or horas[d].sum() > 0]
    z, texto = [], []
    for dep in horas.index:
        vals = [float(horas.loc[dep, d]) for d in dias]
        mx = max(vals) if vals else 0
        z.append([(v / mx * 100 if mx > 0 else None) for v in vals])
        texto.append([f"{v:.1f}" if v > 0 else "" for v in vals])
    fig = go.Figure(go.Heatmap(
        z=z, x=[f"{DIAS_ES[d.weekday()]} {d:%d/%m}" for d in dias], y=[etiqueta_depto(d) for d in horas.index],
        text=texto, texttemplate="%{text}", zmin=0, zmax=100, colorbar=dict(title="% del pico"),
        colorscale=[[0, "#d8f3dc"], [0.5, "#ffe8a3"], [1, "#e63946"]]))
    fig.update_yaxes(autorange="reversed")
    fig.update_xaxes(type="category", tickangle=-60)
    st.plotly_chart(estilo_fig(fig, max(320, 42 * len(horas.index) + 140)))

    seccion("Días de mayor carga por departamento")
    picos = calc.top_dias_pico(horas, cap, 3)
    if not picos.empty:
        picos["Día"] = picos["Día"].map(lambda d: f"{DIAS_ES[d.weekday()]} {d:%d/%m}")
        st.dataframe(picos, hide_index=True)

    seccion("¿Quién puede ayudar un día puntual?")
    ini, fin = calc.rango_mes(anio, mes)
    dia = st.date_input("Día", value=min(max(hoy, ini), fin), min_value=ini, max_value=fin,
                        format="DD/MM/YYYY", key="cal_dia")
    st.dataframe(calc.quien_puede_ayudar(reg, dia, horas_pp, feriados), hide_index=True)


def tab_semanal(ctx, anio, mes):
    reg, feriados, cat = ctx["reg"], ctx["feriados"], ctx["cat"]
    horas_pp = calc.horas_por_persona(ctx["personas"])
    deptos = list(dict.fromkeys(cat["Departamento"]))
    h, res = calc.semanal_departamentos(reg, anio, mes, horas_pp, feriados, deptos)
    if h.to_numpy().sum() == 0:
        st.info("Todavía no hay horas de trabajo cargadas en este mes.")
        return
    st.caption("Cuánto trabajo cae cada semana en cada departamento y cuánto podía absorber el equipo. "
               "**Libre** = capacidad del equipo − trabajo total de la semana: lo que se podría absorber sin horas extra.")
    semanas = list(h.columns)
    fig = go.Figure()
    for dep in h.index:
        if h.loc[dep].sum() > 0:
            fig.add_trace(go.Bar(name=etiqueta_depto(dep), x=semanas, y=h.loc[dep].round(1).tolist(),
                                 marker_color=C.COLORES_DEPTO.get(dep)))
    fig.add_trace(go.Scatter(name="Capacidad del equipo", x=semanas,
                             y=res["Capacidad del equipo (hs)"].tolist(), mode="lines+markers",
                             line=dict(color="#1B2A38", dash="dash")))
    fig.update_layout(barmode="stack", yaxis_title="Horas")
    seccion("Trabajo por semana y departamento")
    st.plotly_chart(estilo_fig(fig, 420))

    seccion("Resumen de la semana")
    st.dataframe(res.reset_index(names="Semana"), hide_index=True)

    seccion("Qué parte de la capacidad del equipo consume cada departamento")
    cap = res["Capacidad del equipo (hs)"]
    z, texto = [], []
    for dep in h.index:
        z.append([(h.loc[dep, s] / cap[s] * 100 if cap[s] > 0 else None) for s in semanas])
        texto.append([f"{h.loc[dep, s]:.1f}" if h.loc[dep, s] > 0 else "" for s in semanas])
    fig2 = go.Figure(go.Heatmap(
        z=z, x=semanas, y=[etiqueta_depto(d) for d in h.index], text=texto, texttemplate="%{text}",
        zmin=0, zmax=60, colorbar=dict(title="% de la capacidad"),
        colorscale=[[0, "#E8F4FB"], [0.5, "#7FB8DA"], [1, "#1F4E78"]]))
    fig2.update_yaxes(autorange="reversed")
    st.plotly_chart(estilo_fig(fig2, max(320, 42 * len(h.index) + 120)))


# ============================================================================
# Manual y protocolo
# ============================================================================
def pantalla_manual():
    st.header("📚 Manual de uso")
    with st.expander("🎯 ¿Qué es este sistema?", expanded=True):
        st.markdown("Registra en qué departamento, tarea y subtarea se usa el tiempo del equipo, para ver la "
                    "capacidad, la carga de cada departamento y decidir si conviene ayudar o contratar.")
    with st.expander("➕ Cargar horas"):
        st.markdown("""
Elegí **Departamento → Tarea → Subtarea**. El departamento es **para qué es el trabajo**, no quién lo hace
(reclamar facturas va en DOCUMENTACIÓN aunque lo haga Atención al cliente).
Los minutos van de a 5, con un mínimo de 10. Para repartir un trabajo en varios días usá *Repartir en varios días*.
""")
    with st.expander("🟦 Disponible, Ausencias y horas extra"):
        st.markdown("""
- **Disponible:** cuando no estás haciendo nada. Elegí la opción 🟦 *Disponible* (no hace falta elegir departamento). Completa tus 6 horas del día.
- **Gestión y mejoras del departamento:** también cuenta como tiempo libre.
- **Ausencias:** elegí 🏖️ *Ausencia* y después *Inasistencia* o *Vacaciones*. Bajan tu capacidad y no cuentan como trabajo. En vacaciones se cargan
  solo los días hábiles entre *Desde* y *Hasta*.
- **Horas extra:** no hay que marcarlas. Si en un día trabajás más que tu capacidad, o trabajás un fin de semana o
  feriado, la diferencia cuenta sola como extra.
""")
    with st.expander("📊 Cómo leer el panel de control"):
        st.markdown("""
- **Capacidad:** días hábiles × 6 hs, menos las ausencias.
- **Utilización:** horas de trabajo sobre la capacidad. **Disponibilidad:** horas libres sobre la capacidad.
- **Desvío:** compara tus horas del mes con el promedio de los 2 meses anteriores. Marca dónde cambió la carga; no es bueno ni malo.
- **Tendencia y semanal:** cómo se reparten tus horas entre departamentos mes a mes y semana a semana.
""")
    with st.expander("🛠️ Para el Admin: calendario y vista semanal"):
        st.markdown("""
- **Calendario:** una fila por departamento y una columna por día. Muestra los días más cargados de cada uno y quién
  puede ayudar un día puntual.
- **Semanal:** cuánto trabajo cae cada semana por departamento y cuántas horas libres tenía el equipo
  (capacidad − trabajo). Sirve para saber si un departamento puede absorber un pico o si conviene sumar gente.
- **Importante:** estos números solo sirven si todos cargan sus 6 horas todos los días (con *Disponible* cuando
  no hacen nada).
""")
    with st.expander("🗂️ Para el Admin: Catálogo, Personas y Feriados"):
        st.markdown("""
Se editan directamente en el Google Sheet:
- **Catalogo:** Departamento, Tarea, Subtarea, Tipo (*Trabajo*, *Disponible* o *Ausencia*), Orden y Activo (SI/NO).
- **Personas:** Nombre, Rol (*Operario* o *Admin*), Activo (SI/NO) y HorasDia. Para sumar a alguien: una fila nueva
  y su contraseña en Secrets, con el mismo nombre.
- **Feriados:** Fecha (AAAA-MM-DD) y Motivo. Hay que agregar los de cada año nuevo.
""")
    with st.expander("✏️ Me equivoqué en una carga"):
        st.markdown("En **Cargar horas → Cargas recientes → Editar o eliminar una carga** podés cambiar los minutos y la "
                    "nota, o borrarla.")


def pantalla_protocolo():
    st.header("📜 Protocolo de uso")
    for titulo, texto in pdf.PROTOCOLO:
        st.markdown(f"**{titulo}**  \n{texto}")
    boton_pdf("protocolo", "📄 Preparar PDF del protocolo", pdf.pdf_protocolo, "Protocolo_Pressacco.pdf")


# ============================================================================
# Programa principal
# ============================================================================
def main():
    try:
        ctx = cargar_contexto()
    except Exception as e:  # noqa: BLE001
        st.error("No pude leer el Google Sheet. Revisá los Secrets y que la cuenta de servicio tenga acceso "
                 f"a «{C.SPREADSHEET}».")
        st.caption(f"Detalle técnico: {e}")
        st.stop()
        return

    if st.session_state.get("usuario") is None:
        pantalla_login(ctx["personas"])
        st.stop()
        return

    usuario = st.session_state["usuario"]
    fila = ctx["personas"][ctx["personas"]["Nombre"] == usuario]
    es_admin = (not fila.empty) and fila["Rol"].iloc[0] == C.ROL_ADMIN

    with st.sidebar:
        st.markdown(f"### 🏛️ Pressacco\n**{usuario}**")
        pagina = st.radio("Navegación", ["➕ Cargar horas", "📊 Panel de control", "📚 Manual", "📜 Protocolo"])
        if st.button("🔄 Actualizar datos"):
            invalidar()
            st.rerun()
        if st.button("Cerrar sesión"):
            st.session_state.clear()
            st.rerun()

    hero(usuario, es_admin, hoy_ar())
    if pagina.startswith("➕"):
        pantalla_carga(ctx, usuario, es_admin)
    elif pagina.startswith("📊"):
        pantalla_resumen(ctx, usuario, es_admin)
    elif pagina.startswith("📚"):
        pantalla_manual()
    else:
        pantalla_protocolo()


if __name__ == "__main__":
    main()
