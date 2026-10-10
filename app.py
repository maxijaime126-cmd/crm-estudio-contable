"""CRM Grupo Pressacco v2 — capacidad instalada por departamento.

Departamento -> Tarea -> Subtarea, minutos, ausencias, horas extra automáticas, calendario de
saturación, vista semanal por departamento, desvío y tendencia, y PDFs.
La lógica está en calc.py (con pruebas), los datos en store.py y los PDFs en pdf.py.
"""
import hmac
import html
import math
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
DURACIONES = {"10 min": 10, "15 min": 15, "30 min": 30, "45 min": 45, "1 h": 60,
              "1 h 30": 90, "2 h": 120, "3 h": 180, "4 h": 240}


def selector_tarea(cat: pd.DataFrame, clave: str, actual: dict | None = None,
                   etiqueta_modo: str = "¿Qué vas a cargar?"):
    """Trabajo > Departamento > Tarea > Subtarea, o Disponible / Ausencia directo.
    'clave' evita choques entre pantallas; 'actual' (opcional) marca lo que viene seleccionado."""
    actual = actual or {}
    modos = ["💼 Trabajo", "🟦 Disponible", "🏖️ Ausencia"]
    i_modo = 2 if actual.get("Tarea") == C.TAREA_AUSENCIAS else 1 if actual.get("Tarea") == C.TAREA_DISPONIBLE else 0
    modo = st.radio(etiqueta_modo, modos, index=i_modo, horizontal=True, key=f"{clave}_modo")
    if modo.startswith("💼"):
        # Disponible y Ausencias tienen su propia opción: acá van las tareas de cada departamento
        cat_t = cat[~cat["Tarea"].isin([C.TAREA_DISPONIBLE, C.TAREA_AUSENCIAS])]
        if cat_t.empty:
            st.error("El catálogo no tiene tareas de trabajo. Revisá la hoja Catalogo.")
            return None
        deptos = list(dict.fromkeys(cat_t["Departamento"]))
        i_dep = deptos.index(actual["Departamento"]) if actual.get("Departamento") in deptos else 0
        depto = st.selectbox("Departamento", deptos, index=i_dep, format_func=etiqueta_depto, key=f"{clave}_dep")
        sub_d = cat_t[cat_t["Departamento"] == depto]
        tareas = list(dict.fromkeys(sub_d["Tarea"]))
        mismo_dep = actual.get("Departamento") == depto
        i_tar = tareas.index(actual["Tarea"]) if mismo_dep and actual.get("Tarea") in tareas else 0
        c2, c3 = st.columns(2)
        tarea = c2.selectbox("Tarea", tareas, index=i_tar, key=f"{clave}_tar_{depto}")
        subs = [x for x in sub_d[sub_d["Tarea"] == tarea]["Subtarea"] if x]
        if subs:
            misma = mismo_dep and actual.get("Tarea") == tarea
            i_sub = subs.index(actual["Subtarea"]) if misma and actual.get("Subtarea") in subs else 0
            subtarea = c3.selectbox("Subtarea", subs, index=i_sub, key=f"{clave}_sub_{depto}_{tarea}")
        else:
            subtarea = ""
            c3.caption("Esta tarea no tiene subtareas.")
        fila_cat = sub_d[(sub_d["Tarea"] == tarea) & (sub_d["Subtarea"] == subtarea)]
        tipo = fila_cat["Tipo"].iloc[0] if not fila_cat.empty else C.TIPO_TRABAJO
        ruta = f"<b>{html.escape(depto)}</b> › {html.escape(tarea)}" + (f" › {html.escape(subtarea)}" if subtarea else "")
        return {"depto": depto, "tarea": tarea, "subtarea": subtarea, "tipo": tipo,
                "color": C.COLORES_DEPTO.get(depto, "#0077B6"), "icono": C.ICONOS_DEPTO.get(depto, "📁"), "ruta": ruta}
    if modo.startswith("🟦"):
        st.caption("Para cuando no estás haciendo nada: completa tus horas del día. No hace falta elegir departamento.")
        return {"depto": C.DEPTO_GENERAL, "tarea": C.TAREA_DISPONIBLE, "subtarea": "", "tipo": C.TIPO_DISPONIBLE,
                "color": "#0096C7", "icono": "🟦", "ruta": "<b>Disponible</b>"}
    opciones = list(dict.fromkeys(x for x in cat[cat["Tarea"] == C.TAREA_AUSENCIAS]["Subtarea"] if x)) \
        or ["Inasistencia (día personal)", "Vacaciones"]
    i_aus = opciones.index(actual["Subtarea"]) if actual.get("Subtarea") in opciones else 0
    subtarea = st.selectbox("Tipo de ausencia", opciones, index=i_aus, key=f"{clave}_aus")
    return {"depto": C.DEPTO_GENERAL, "tarea": C.TAREA_AUSENCIAS, "subtarea": subtarea, "tipo": C.TIPO_AUSENCIA,
            "color": "#6B7C8C", "icono": "🏖️", "ruta": f"<b>Ausencia</b> › {html.escape(subtarea)}"}


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
    sel = selector_tarea(cat, "c")
    if sel is None:
        return
    depto, tarea, subtarea, tipo = sel["depto"], sel["tarea"], sel["subtarea"], sel["tipo"]
    es_ausencia = tipo == C.TIPO_AUSENCIA
    color = sel["color"]
    st.markdown(
        f'<div class="dep-card" style="border-left:8px solid {color}; background:{color}22;">'
        f'{sel["icono"]} {sel["ruta"]}<span class="chip {CHIP_TIPO.get(tipo, "chip-trab")}">'
        f'{html.escape(tipo)}</span></div>', unsafe_allow_html=True)

    # --- Cuándo y cuánto ---
    ver = st.session_state.get("f_ver", 0)   # sube al guardar: así los campos vuelven a su valor inicial
    if es_ausencia:
        st.caption("Las ausencias bajan tu capacidad del mes; no cuentan como trabajo.")
        d1, d2 = st.columns(2)
        desde = d1.date_input("Desde", value=hoy, format="DD/MM/YYYY", key=f"a_desde_{ver}")
        hasta = d2.date_input("Hasta", value=hoy, format="DD/MM/YYYY", key=f"a_hasta_{ver}")
        completo = st.checkbox(f"Día completo ({fmt_hs(hd)} hs por día)", value=True, key=f"a_comp_{ver}")
        if completo:
            minutos = int(hd * 60)
        else:
            minutos = st.number_input("Minutos por día", min_value=C.MIN_MINUTOS,
                                      max_value=int(hd * 60), step=C.PASO_MINUTOS, value=120, key=f"a_min_{ver}")
        modo, solo_habiles = "por_dia", True
    else:
        modalidad = st.radio("¿Cuándo?", ["Un día", "Repartir en varios días"], horizontal=True)
        if modalidad == "Un día":
            desde = hasta = st.date_input("Fecha", value=hoy, max_value=hoy, format="DD/MM/YYYY", key="t_fecha")
            dur = st.radio("Duración", list(DURACIONES) + ["Otra"], index=4, horizontal=True, key=f"t_dur_{ver}")
            if dur == "Otra":
                minutos = st.number_input("Minutos", min_value=C.MIN_MINUTOS, max_value=C.MAX_MINUTOS_CARGA,
                                          step=C.PASO_MINUTOS, value=60, key=f"t_min_{ver}")
            else:
                minutos = DURACIONES[dur]
            modo = "por_dia"
        else:
            d1, d2 = st.columns(2)
            desde = d1.date_input("Desde", value=hoy, max_value=hoy, format="DD/MM/YYYY", key=f"t_desde_{ver}")
            hasta = d2.date_input("Hasta", value=hoy, max_value=hoy, format="DD/MM/YYYY", key=f"t_hasta_{ver}")
            minutos = st.number_input("Minutos totales a repartir", min_value=C.MIN_MINUTOS,
                                      max_value=60 * 200, step=C.PASO_MINUTOS, value=600, key=f"t_tot_{ver}")
            modo = "total"
        solo_habiles = False
        st.caption(f"= {fmt_min(minutos)}")

    obligatoria = subtarea == C.SUBTAREA_NOTA_OBLIGATORIA
    nota = st.text_input("Nota (obligatoria: contá qué fue)" if obligatoria else "Nota (opcional)", key=f"c_nota_{ver}")

    # --- Vista previa ---
    filas = []
    resto = 0   # minutos que faltarían para completar el día, después de esta carga
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
        if len(filas) == 1:
            resto = resumen_dia(reg, persona, desde, hd, feriados, filas[0]["Minutos"] / 60, tipo)

    completar = False
    if tipo == C.TIPO_TRABAJO and len(filas) == 1 and resto >= C.PASO_MINUTOS:
        st.caption(f"Si el resto del día estuvo libre, se cargarían {fmt_min(resto)} de Disponible.")
        completar = st.checkbox("🟦 Completar el día con Disponible al guardar", key=f"c_compl_{ver}")

    if st.button("💾 Guardar", type="primary", disabled=not filas):
        if obligatoria and not nota.strip():
            st.error("Para «Otros imprevistos» la nota es obligatoria.")
        else:
            guardado, error = False, None
            if completar and resto >= C.PASO_MINUTOS:
                filas = filas + [calc.registro_disponible(persona, desde, resto)]
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
                st.session_state["f_ver"] = ver + 1   # limpia minutos y nota para la próxima carga
                st.rerun()
            else:
                st.error(f"No se pudo guardar: {error}")

    seccion_mis_cargas(reg, persona, hoy, hd, feriados, desde, cat)


def resumen_dia(reg, persona, dia, hd, feriados, nuevas_hs, tipo):
    """Cuánto lleva cargado ese día y cómo queda con la carga que está por guardar."""
    rp = reg[(reg["Persona"] == persona) & (reg["Fecha"].dt.date == dia)]
    d = calc.por_dia(rp, hd, feriados, dia, dia).iloc[0]
    llevas = float(d[C.TIPO_TRABAJO] + d[C.TIPO_DISPONIBLE] + d[C.TIPO_AUSENCIA])
    meta = hd if calc.es_habil(dia, feriados) else 0.0
    total = llevas + nuevas_hs
    de_meta = f" de {fmt_hs(meta)} hs" if meta else ""
    st.markdown(
        f'<div class="obj-box">📅 <b>{dia:%d/%m}</b> · llevás <b>{fmt_hs(llevas)} hs</b> cargadas{de_meta}'
        f' · con esta carga: <b>{fmt_hs(total)} hs</b></div>', unsafe_allow_html=True)
    if meta > 0:
        st.progress(min(1.0, total / meta))
        falta = meta - total
        if falta > 0.03:
            st.caption(f"Con esta carga todavía te faltarían {fmt_hs(falta)} hs para completar el día.")
        elif falta > -0.03:
            st.caption("✅ Con esta carga completás las horas del día.")
    else:
        st.caption("Día no hábil (fin de semana o feriado): lo que trabajes cuenta como horas extra.")
    if tipo == C.TIPO_TRABAJO:
        extra = max(0.0, float(d[C.TIPO_TRABAJO]) + nuevas_hs - float(d["Capacidad"]))
        if extra > 0.03:
            st.warning(f"Con esta carga ese día quedan {fmt_hs(extra)} hs extra.")
    return max(0, round((meta - total) * 60)) if meta > 0 else 0


def tabla_cargas(df: pd.DataFrame, con_fecha: bool) -> pd.DataFrame:
    """Tabla para mostrar: duración en horas y minutos, y una fila final con el TOTAL."""
    v = pd.DataFrame({
        "Departamento": df["Departamento"].tolist(), "Tarea": df["Tarea"].tolist(),
        "Subtarea": df["Subtarea"].tolist(), "Duración": [fmt_min(m) for m in df["Minutos"]],
        "Minutos": df["Minutos"].astype(int).tolist(), "Nota": df["Nota"].tolist()})
    if con_fecha:
        v.insert(0, "Fecha", df["Fecha"].dt.strftime("%d/%m/%Y").tolist())
    total = int(df["Minutos"].sum())
    fila = {c: "" for c in v.columns}
    fila.update({"Departamento": "TOTAL", "Duración": fmt_min(total), "Minutos": total})
    return pd.concat([v, pd.DataFrame([fila])], ignore_index=True)


def boton_completar(persona: str, dia: date, hd: float, feriados: set, hoy: date, total_min: int):
    """Carga de una vez como Disponible las horas que faltan para completar el día."""
    meta_min = hd * 60 if calc.es_habil(dia, feriados) else 0
    faltan = round(meta_min - total_min)
    if faltan < C.PASO_MINUTOS or dia > hoy:
        return
    if total_min == 0:
        st.caption(f"Faltan {fmt_min(faltan)} (el día está sin cargas). Si estuviste libre todo el día, podés cargarlo de una vez.")
    else:
        st.caption(f"Faltan {fmt_min(faltan)} para completar el día. Si el resto estuvo libre, podés cargarlo de una vez.")
    if st.button("🟦 Completar el día con Disponible", key=f"mc_compl_{persona}_{dia}"):
        error = None
        try:
            fila = calc.registro_disponible(persona, dia, faltan)
            fila.update({"ID": S.nuevo_id(), "Registrado": S.ahora_iso()})
            get_store().agregar(C.HOJA_REGISTROS, [fila])
            invalidar()
        except Exception as e:  # noqa: BLE001
            error = e
        if error:
            st.error(f"No se pudo guardar: {error}")
        else:
            st.session_state["msg_ok"] = f"✅ Se cargaron {fmt_min(faltan)} de Disponible el {dia:%d/%m/%Y}."
            st.rerun()


def seccion_mis_cargas(reg: pd.DataFrame, persona: str, hoy: date, hd: float, feriados: set, dia_ref: date,
                       cat: pd.DataFrame):
    st.divider()
    seccion("🗂️ Mis cargas")
    propias = reg[reg["Persona"] == persona]
    ver = st.radio("Ver", ["📅 Del día", "🗓️ Del mes"], horizontal=True, key="mc_ver")
    if ver.startswith("📅"):
        dia = st.date_input("Día", value=dia_ref, format="DD/MM/YYYY")
        mostradas = propias[propias["Fecha"].dt.date == dia].sort_values("Registrado")
        if mostradas.empty:
            st.caption(f"No hay cargas de {persona} el {dia:%d/%m/%Y}.")
            boton_completar(persona, dia, hd, feriados, hoy, 0)
            return
        por_tipo = mostradas.groupby("Tipo")["Horas"].sum()
        total_min = int(mostradas["Minutos"].sum())
        st.markdown(
            f'<div class="obj-box">📅 <b>{dia:%d/%m/%Y}</b> · total <b>{fmt_min(total_min)}</b> '
            f'({fmt_hs(total_min / 60)} hs) · trabajo {fmt_hs(float(por_tipo.get(C.TIPO_TRABAJO, 0)))} hs · '
            f'disponible {fmt_hs(float(por_tipo.get(C.TIPO_DISPONIBLE, 0)))} hs · '
            f'ausencia {fmt_hs(float(por_tipo.get(C.TIPO_AUSENCIA, 0)))} hs</div>', unsafe_allow_html=True)
        st.dataframe(tabla_cargas(mostradas, False), hide_index=True)
        boton_completar(persona, dia, hd, feriados, hoy, total_min)
    else:
        anios = sorted({hoy.year, *propias["Fecha"].dt.year.astype(int).tolist()})
        f1, f2, _ = st.columns([1.3, 1.6, 2.5])
        anio = f1.selectbox("Año", anios, index=anios.index(hoy.year), key="mc_anio")
        mes = f2.selectbox("Mes", list(range(1, 13)), index=hoy.month - 1,
                           format_func=lambda m: C.MESES_ES[m], key="mc_mes")
        mostradas = calc.del_mes(propias, anio, mes).sort_values(["Fecha", "Registrado"], ascending=False)
        if mostradas.empty:
            st.caption(f"No hay cargas de {persona} en {C.MESES_ES[mes]} {anio}.")
            return
        total_min = int(mostradas["Minutos"].sum())
        st.caption(f"{len(mostradas)} carga(s) · {fmt_min(total_min)} ({fmt_hs(total_min / 60)} hs) "
                   f"en {C.MESES_ES[mes]} {anio}")
        resumen = calc.resumen_diario(reg, persona, anio, mes, hd, feriados, hoy)
        if not resumen.empty:
            st.markdown("**Día por día** (¿quedó completo?)")
            resumen = resumen.copy()
            resumen["Fecha"] = resumen["Fecha"].map(lambda d: f"{DIAS_ES[d.weekday()]} {d:%d/%m}")
            st.dataframe(resumen, hide_index=True)
        with st.expander("Ver todas las cargas del mes"):
            st.dataframe(tabla_cargas(mostradas, True), hide_index=True)
    editar_eliminar(mostradas, cat)


def editar_eliminar(mostradas: pd.DataFrame, cat: pd.DataFrame):
    editables = mostradas[mostradas["ID"] != ""]
    if editables.empty:
        return
    with st.expander("✏️ Editar o eliminar una carga"):
        etiquetas = {r.ID: (f"{r.Fecha:%d/%m} · {r.Departamento} › {r.Tarea}"
                            f"{' › ' + r.Subtarea if r.Subtarea else ''} · {fmt_min(int(r.Minutos))}")
                     for r in editables.itertuples()}
        elegido = st.selectbox("Carga", list(etiquetas), format_func=lambda i: etiquetas[i])
        fila = editables[editables["ID"] == elegido].iloc[0]

        fecha_nueva = st.date_input("Fecha", value=fila["Fecha"].date(), format="DD/MM/YYYY",
                                    key=f"e_fecha_{elegido}")
        valor = max(C.PASO_MINUTOS, min(int(fila["Minutos"]), C.MAX_MINUTOS_CARGA))
        nuevos = st.number_input("Minutos", min_value=C.PASO_MINUTOS, max_value=C.MAX_MINUTOS_CARGA,
                                 step=C.PASO_MINUTOS, value=valor, key=f"e_min_{elegido}")
        st.caption(f"= {fmt_min(nuevos)}")
        nueva_nota = st.text_input("Nota", value=fila["Nota"], key=f"e_nota_{elegido}")
        sel = None
        if st.checkbox("Cambiar departamento / tarea / subtarea", key=f"e_cambiar_{elegido}"):
            sel = selector_tarea(cat, f"e{elegido}", {"Departamento": fila["Departamento"], "Tarea": fila["Tarea"],
                                                       "Subtarea": fila["Subtarea"]},
                                 etiqueta_modo="¿Qué tipo de carga es?")
        b1, b2 = st.columns(2)
        if b1.button("Guardar cambios", key=f"e_ok_{elegido}"):
            subtarea_final = sel["subtarea"] if sel else fila["Subtarea"]
            if subtarea_final == C.SUBTAREA_NOTA_OBLIGATORIA and not nueva_nota.strip():
                st.error("Para «Otros imprevistos» la nota es obligatoria.")
            else:
                cambios = {"Fecha": fecha_nueva.isoformat(), "Minutos": int(nuevos), "Nota": nueva_nota.strip()}
                if sel:
                    cambios.update({"Departamento": sel["depto"], "Tarea": sel["tarea"], "Subtarea": sel["subtarea"]})
                try:
                    get_store().actualizar_por_id(C.HOJA_REGISTROS, elegido, cambios)
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
def banner_pendientes(ctx: dict, usuario: str, hoy: date):
    """Aviso para quien carga horas: días hábiles ya pasados que no completó."""
    horas_pp = calc.horas_por_persona(ctx["personas"])
    if usuario not in horas_pp:
        return
    meta = horas_pp[usuario] * 60
    pend = calc.dias_pendientes(ctx["reg"], usuario, horas_pp[usuario], ctx["feriados"], hoy)
    if not pend:
        return
    items = [f"{d:%d/%m} ({'sin carga' if m >= meta - 2 else 'faltan ' + fmt_min(m)})" for d, m in pend[:6]]
    texto = (f"📅 Tenés {len(pend)} día(s) sin completar: " + ", ".join(items) + (" y más" if len(pend) > 6 else "")
             + ". Cargalos desde «Cargar horas» (en «Mis cargas» podés completar un día con Disponible).")
    alertas([texto])


def bloque_estado_dia(ctx: dict, hoy: date):
    """Para el Admin: ¿quién completó las horas del día?"""
    horas_pp = calc.horas_por_persona(ctx["personas"])
    if not horas_pp:
        return
    seccion("🕒 ¿Quién completó el día?")
    dia = st.date_input("Día a revisar", value=hoy, max_value=hoy, format="DD/MM/YYYY", key="estado_dia")
    t = calc.estado_dia(ctx["reg"], dia, horas_pp, ctx["feriados"])
    if not calc.es_habil(dia, ctx["feriados"]):
        st.caption("Ese día no es hábil (fin de semana o feriado).")
    else:
        completos = int(t["Estado"].str.startswith("✅").sum())
        st.markdown(f'<div class="obj-box">✅ <b>{completos} de {len(t)}</b> completaron las horas del '
                    f'{dia:%d/%m/%Y}</div>', unsafe_allow_html=True)
    st.dataframe(t, hide_index=True)
    if dia == hoy:
        st.caption("El protocolo pide cargar las horas antes de las 15 hs: antes de esa hora es normal que falten.")


def pantalla_resumen(ctx: dict, usuario: str, es_admin: bool):
    reg = ctx["reg"]
    hoy = hoy_ar()
    st.header("📊 Panel de control")
    if es_admin:
        bloque_estado_dia(ctx, hoy)
        st.divider()
    anios = sorted({hoy.year - 1, hoy.year, hoy.year + 1, *reg["Fecha"].dt.year.astype(int).tolist()})
    f1, f2, _ = st.columns([1.3, 1.6, 2.5])
    anio = f1.selectbox("Año", anios, index=anios.index(hoy.year))
    mes = f2.selectbox("Mes", list(range(1, 13)), index=hoy.month - 1, format_func=lambda m: C.MESES_ES[m])

    if es_admin:
        t1, t2, t5, t3, t4 = st.tabs(["👤 Individual", "🌐 Equipo", "🏢 Departamentos", "🗓️ Calendario", "📆 Semanal"])
        with t1:
            tab_individual(ctx, None, anio, mes, hoy, es_admin)
        with t2:
            tab_equipo(ctx, anio, mes)
        with t5:
            tab_departamentos(ctx, anio, mes)
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


def seccion_semanal_personas(ctx, anio, mes):
    """Semana por semana, persona por persona: quién está ocupado, en qué departamento y quién tiene margen."""
    reg, feriados = ctx["reg"], ctx["feriados"]
    horas_pp = calc.horas_por_persona(ctx["personas"])
    personas = list(horas_pp)
    sp = calc.semanal_por_persona(reg, anio, mes, horas_pp, feriados, hoy_ar())
    dp, semanas = calc.semanal_persona_departamento(reg, anio, mes, horas_pp)
    seccion("Distribución semanal por persona")
    if sp.empty or float(sp["Capacidad (hs)"].sum()) == 0:
        st.info("Todavía no hay horas cargadas en este mes.")
        return
    st.caption("Cada persona, semana por semana: cuánto trabajó, cuánto margen tuvo y en qué departamento. Solo cuentan "
               "los días en que esa persona cargó horas (los días sin cargar no se toman como libres). "
               "**Libre** = capacidad − trabajo.")
    z, texto = [], []
    for p in personas:
        fz, ft = [], []
        for sem in semanas:
            r = sp[(sp["Persona"] == p) & (sp["Semana"] == sem)].iloc[0]
            fz.append(float(r["Ocupación %"]) if r["Capacidad (hs)"] > 0 else float("nan"))
            ft.append(f"{r['Trabajo (hs)']:.1f} hs<br>libre {r['Libre (hs)']:.1f}" if r["Capacidad (hs)"] > 0 else "")
        z.append(fz)
        texto.append(ft)
    fig = go.Figure(go.Heatmap(
        z=z, x=semanas, y=personas, text=texto, texttemplate="%{text}", zmin=0, zmax=100, colorscale=ESCALA_CARGA,
        colorbar=dict(title="% ocupado"), hoverongaps=False, xgap=3, ygap=3))
    fig.update_yaxes(autorange="reversed")
    estilo_fig(fig, max(240, 80 * len(personas) + 100))
    fig.update_layout(plot_bgcolor=GRIS_SIN_DATOS)
    st.plotly_chart(fig)

    if not dp.empty:
        st.markdown("**¿En qué departamentos trabajó cada persona?**")
        dp = dp.copy()
        cortas = {x: x.split(" (")[0] for x in semanas}
        dp["Semana"] = dp["Semana"].map(cortas)
        fig2 = px.bar(dp, x="Semana", y="Horas", color="Departamento", facet_col="Persona", facet_col_wrap=2,
                      color_discrete_map=C.COLORES_DEPTO,
                      category_orders={"Semana": list(cortas.values()), "Persona": personas})
        fig2.for_each_annotation(lambda a: a.update(text=a.text.split("=")[-1]))
        fig2.update_xaxes(title=None)
        st.plotly_chart(estilo_fig(fig2, 270 * math.ceil(len(personas) / 2) + 40))

    st.markdown("**¿Quién tiene margen para ayudar?**")
    con_datos = [x for x in semanas if float(sp[sp["Semana"] == x]["Capacidad (hs)"].sum()) > 0]
    sel = st.selectbox("Semana", con_datos, index=len(con_datos) - 1, key="sp_sem")
    t = sp[sp["Semana"] == sel].sort_values(["Libre (hs)", "Persona"], ascending=[False, True])
    top = t.iloc[0]
    if float(top["Libre (hs)"]) > 0:
        donde = f", sobre todo en {top['Departamento principal']}" if top["Departamento principal"] else ""
        st.info(f"Con más margen en esa semana: {top['Persona']} ({fmt_hs(float(top['Libre (hs)']))} hs libres{donde}).")
    cols = ["Persona", "Trabajo (hs)", "Disponible (hs)", "Libre (hs)", "Ocupación %", "Departamento principal",
            "% en ese departamento", "Días sin carga"]
    st.dataframe(t[cols], hide_index=True)


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

    seccion_semanal_personas(ctx, anio, mes)

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


DIAS_LARGO = ["Lun", "Mar", "Mié", "Jue", "Vie", "Sáb", "Dom"]
ESCALA_CARGA = [[0, "#d8f3dc"], [0.5, "#ffe8a3"], [1, "#e63946"]]
GRIS_SIN_DATOS = "#E3E7EC"


def tab_departamentos(ctx, anio, mes):
    """Cuántas horas consume cada departamento, sus picos y si el equipo puede absorberlos."""
    reg, feriados, cat = ctx["reg"], ctx["feriados"], ctx["cat"]
    horas_pp = calc.horas_por_persona(ctx["personas"])
    deptos = list(dict.fromkeys(cat["Departamento"]))
    if not horas_pp:
        st.warning("No hay operarios activos en la hoja Personas.")
        return
    res = calc.resumen_departamentos_mes(reg, anio, mes, horas_pp, feriados, deptos)
    if res["Horas"].sum() == 0:
        st.info("Todavía no hay horas de trabajo cargadas en este mes.")
        return
    periodo = f"{C.MESES_ES[mes]} {anio}"

    seccion(f"Horas de trabajo por departamento — {periodo}")
    top = res.iloc[0]
    kpis([(fmt_hs(float(res["Horas"].sum())), "Trabajo del equipo (hs)", "k-trab"),
          (f"{fmt_hs(float(top['Horas']))} hs", f"Más cargado: {top['Departamento']}", "k-cap"),
          (f"{float(res['% de la capacidad del equipo'].sum()):.0f}%", "De la capacidad del equipo", "k-util")])
    graf = res[res["Horas"] > 0].sort_values("Horas")
    fig = px.bar(graf, x="Horas", y="Departamento", orientation="h", color="Departamento",
                 color_discrete_map=C.COLORES_DEPTO, text=graf["Horas"].round(1))
    fig.update_layout(showlegend=False, yaxis_title=None)
    st.plotly_chart(estilo_fig(fig, max(260, 40 * len(graf) + 80)))
    st.caption("**Personas equivalentes** = horas del mes ÷ (días hábiles × horas por día): cuántas personas a tiempo "
               "completo hicieron falta para ese departamento en el mes.")
    tabla = res[(res["Horas"] > 0) | (res["Mes anterior (hs)"] > 0)].copy()
    st.dataframe(tabla, hide_index=True)

    seccion("Ventana más cargada de cada departamento")
    n = st.radio("Ventana de", [3, 4, 5], index=1, horizontal=True, key="vent_n",
                 format_func=lambda x: f"{x} días hábiles seguidos")
    v = calc.ventanas_pico(reg, anio, mes, horas_pp, feriados, deptos, n)
    if v.empty:
        st.info(f"Todavía no hay {n} días hábiles seguidos con horas cargadas en este mes.")
    else:
        st.caption(f"Los {n} días hábiles seguidos de mayor carga de cada departamento. **Una persona alcanza** = lo que "
                   f"una sola persona puede hacer en esos días ({n} × 6 hs). **Libre del equipo** = capacidad de todos − "
                   "trabajo de todos en esos días. Solo se miran ventanas donde se cargó todos los días.")
        vv = v.copy()
        vv["Desde"] = vv["Desde"].map(lambda d: f"{DIAS_ES[d.weekday()]} {d:%d/%m}")
        vv["Hasta"] = vv["Hasta"].map(lambda d: f"{DIAS_ES[d.weekday()]} {d:%d/%m}")
        st.dataframe(vv, hide_index=True)

    seccion("📄 Informe para la reunión con el contador")
    st.caption("Un PDF con las horas por departamento, los días saturados, el tiempo disponible y lo más importante "
               "del mes, listo para mostrar. Si el mes está en curso, los números llegan hasta el último día cargado.")
    boton_pdf(f"cont_{anio}_{mes}_{n}", "📄 Preparar informe para el contador",
              lambda: pdf.pdf_departamentos(calc.informe_departamentos(
                  reg, anio, mes, horas_pp, feriados, deptos, hoy_ar(), n)),
              f"Informe_departamentos_{anio}-{mes:02d}.pdf")

    seccion("🧮 ¿Podemos absorber un pico?")
    st.caption(f"Por ejemplo: «un departamento necesita 20 horas en 4 días». Se compara con las horas libres que tuvo "
               f"el equipo en {periodo} (promedio por día hábil con datos).")
    c1, c2 = st.columns(2)
    horas_nec = c1.number_input("Horas que necesita el departamento", min_value=1, max_value=500, value=20, step=1,
                                key="abs_horas")
    dias_nec = c2.number_input("En cuántos días hábiles", min_value=1, max_value=20, value=4, step=1, key="abs_dias")
    r = calc.puede_absorber(reg, anio, mes, horas_pp, feriados, float(horas_nec), int(dias_nec))
    if r is None:
        st.info("Todavía no hay días con horas cargadas en este mes para usar como referencia.")
        return
    pct = f" · {r['pct_capacidad']:.0f}% de la capacidad del equipo" if r["pct_capacidad"] is not None else ""
    st.markdown(
        f'<div class="obj-box">🧮 <b>{horas_nec} hs en {dias_nec} días</b> = {r["personas_equivalentes"]:.2f} personas '
        f'equivalentes{pct}<br>El equipo tuvo en promedio <b>{fmt_hs(r["libre_por_dia"])} hs libres por día</b> '
        f'→ en {dias_nec} días: <b>{fmt_hs(r["libre_ventana"])} hs</b> (referencia: {r["dias_usados"]} día(s) con datos)</div>',
        unsafe_allow_html=True)
    if r["alcanza"]:
        st.success("✅ Entre todos se puede absorber: las horas libres alcanzan.")
    else:
        st.warning(f"⚠️ No alcanza solo con las horas libres: faltarían {fmt_hs(r['faltan'])} hs. Habría que "
                   "redistribuir tareas, correr plazos o contar con horas extra.")


def tab_calendario(ctx, anio, mes, hoy):
    reg, feriados, cat = ctx["reg"], ctx["feriados"], ctx["cat"]
    horas_pp = calc.horas_por_persona(ctx["personas"])
    deptos = list(dict.fromkeys(cat["Departamento"]))
    horas, cap = calc.matriz_departamentos(reg, anio, mes, horas_pp, feriados, deptos)
    if horas.to_numpy().sum() == 0:
        st.info("Todavía no hay horas de trabajo cargadas en este mes.")
        return
    con_datos = calc.dias_con_datos(reg, anio, mes)

    o1, o2 = st.columns(2)
    modo_txt = o1.radio("Colorear según", ["Capacidad del equipo", "Pico de cada departamento"], key="cal_modo")
    vista = o2.radio("Vista", ["Mes completo", "Calendario de un departamento"], key="cal_vista")
    modo = "equipo" if modo_txt.startswith("Capacidad") else "pico"
    if modo == "equipo":
        st.caption("El color muestra qué parte de la capacidad del equipo de ese día se usó en el departamento "
                   "(rojo = el equipo entero ocupado en ese departamento). **Gris = nadie cargó horas ese día**: "
                   "no se sabe si estuvo libre o si faltó cargar.")
        titulo = "% de la capacidad del equipo"
    else:
        st.caption("El color compara los días de un mismo departamento (rojo = su día de mayor carga del mes). "
                   "Sirve para ver los días fuertes de un departamento chico. Gris = sin datos.")
        titulo = "% del día pico"

    if vista == "Mes completo":
        z, texto = calc.calendario_z(horas, cap, con_datos, modo)
        dias = [d for d in z.columns if calc.es_habil(d, feriados) or horas[d].sum() > 0]
        filas = [("👥 " + calc.TOTAL_EQUIPO) if i == calc.TOTAL_EQUIPO else etiqueta_depto(i) for i in z.index]
        fig = go.Figure(go.Heatmap(
            z=z[dias].values.tolist(), x=[f"{DIAS_ES[d.weekday()]} {d:%d/%m}" for d in dias], y=filas,
            text=texto[dias].values.tolist(), texttemplate="%{text}", zmin=0, zmax=100,
            colorscale=ESCALA_CARGA, colorbar=dict(title=titulo), hoverongaps=False, xgap=2, ygap=2))
        fig.update_yaxes(autorange="reversed")
        fig.update_xaxes(type="category", tickangle=-60)
        estilo_fig(fig, max(340, 42 * len(filas) + 150))
        fig.update_layout(plot_bgcolor=GRIS_SIN_DATOS)
        st.plotly_chart(fig)
    else:
        dep = st.selectbox("Departamento", deptos, format_func=etiqueta_depto, key="cal_dep")
        semanas, z, texto = calc.grilla_calendario(horas.loc[dep], cap, con_datos, feriados, anio, mes, modo)
        fin_de_semana = sum(float(horas.loc[dep, d]) for d in horas.columns if d.weekday() >= 5) > 0
        n = 7 if fin_de_semana else 5
        fig = go.Figure(go.Heatmap(
            z=[f[:n] for f in z], x=DIAS_LARGO[:n], y=semanas, text=[f[:n] for f in texto],
            texttemplate="%{text}", textfont=dict(size=14), zmin=0, zmax=100, colorscale=ESCALA_CARGA,
            colorbar=dict(title=titulo), hoverongaps=False, xgap=3, ygap=3))
        fig.update_yaxes(autorange="reversed")
        fig.update_xaxes(side="top")
        estilo_fig(fig, max(280, 100 * len(semanas) + 90))
        fig.update_layout(plot_bgcolor=GRIS_SIN_DATOS, margin=dict(l=0, r=0, t=60, b=0))
        st.plotly_chart(fig)
        st.caption(f"{etiqueta_depto(dep)}: {fmt_hs(float(horas.loc[dep].sum()))} hs de trabajo en "
                   f"{C.MESES_ES[mes]} {anio}.")

    seccion("Días de mayor carga por departamento")
    picos = calc.top_dias_pico(horas, cap, 3)
    if not picos.empty:
        picos["Día"] = picos["Día"].map(lambda d: f"{DIAS_ES[d.weekday()]} {d:%d/%m}")
        st.dataframe(picos, hide_index=True)

    seccion("¿Quién puede ayudar un día puntual?")
    ini, fin = calc.rango_mes(anio, mes)
    dia = st.date_input("Día", value=min(max(hoy, ini), fin), min_value=ini, max_value=fin,
                        format="DD/MM/YYYY", key="cal_dia")
    ayuda = calc.quien_puede_ayudar(reg, dia, horas_pp, feriados)
    if not ayuda.empty and (ayuda["Estado"] == "con carga").sum() == 0:
        st.info("Nadie cargó horas ese día (o no es un día hábil): no se puede saber quién tenía margen.")
    else:
        st.caption("Margen = capacidad del día − trabajo cargado. Si alguien no cargó nada ese día figura "
                   "«sin carga»: no se asume que estaba libre.")
        st.dataframe(ayuda, hide_index=True)


def tab_semanal(ctx, anio, mes):
    reg, feriados, cat = ctx["reg"], ctx["feriados"], ctx["cat"]
    horas_pp = calc.horas_por_persona(ctx["personas"])
    deptos = list(dict.fromkeys(cat["Departamento"]))
    h, res = calc.semanal_departamentos(reg, anio, mes, horas_pp, feriados, deptos)
    if h.to_numpy().sum() == 0:
        st.info("Todavía no hay horas de trabajo cargadas en este mes.")
        return
    st.caption("Cuánto trabajo cae cada semana en cada departamento y cuánto podía absorber el equipo. "
               "**Libre** = capacidad del equipo − trabajo total de la semana: lo que se podría absorber sin horas extra. "
               "Solo se cuentan los días en que alguien cargó horas, así las semanas que todavía no pasaron no figuran como libres.")
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
La duración se elige con un toque (10 min, 15, 30, 45, 1 h, 1 h 30, 2 h, 3 h, 4 h) o con *Otra* para escribir los minutos exactos (de a 5, mínimo 10). Para repartir un trabajo en varios días usá *Repartir en varios días*.
""")
    with st.expander("🟦 Disponible, Ausencias y horas extra"):
        st.markdown("""
- **Disponible:** cuando no estás haciendo nada. Elegí la opción 🟦 *Disponible* (no hace falta elegir departamento). Completa tus 6 horas del día.
- **Completar el día:** al cargar trabajo podés tildar *Completar el día con Disponible al guardar*, o usar el botón en *Mis cargas → Del día*: carga de una vez las horas que faltan.
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
- **Calendario:** una fila por departamento y una columna por día, con una fila *Total equipo*. Se puede colorear según la
  capacidad del equipo o según el pico de cada departamento, y ver un calendario de semanas de un departamento.
  El gris significa que nadie cargó horas ese día. Abajo, los días más cargados y quién puede ayudar un día puntual.
- **Departamentos:** cuántas horas consume cada departamento en el mes (contra el mes anterior), la ventana de 3, 4 o 5 días
  más cargada de cada uno, un **PDF para la reunión con el contador** y una calculadora: «20 horas en 4 días, ¿alcanza con las horas libres del equipo?».
- **Equipo → Distribución semanal por persona:** semana por semana, cuánto trabajó cada persona, cuánto margen tuvo y en qué
  departamentos, para ver quién puede ayudar a quién.
- **Quién completó el día:** arriba del Panel de control, el Admin ve quién cargó las horas del día (por defecto hoy).
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
        st.markdown("En **Cargar horas → Mis cargas → Editar o eliminar una carga** podés cambiar la fecha, los minutos, la "
                    "nota y, tildando la opción, el departamento, la tarea y la subtarea. También podés borrarla. En *Mis cargas* ves lo del día (con el total) o el mes día por día.")


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
        pagina = st.radio("Navegación", ["📊 Panel de control", "➕ Cargar horas", "📚 Manual", "📜 Protocolo"])
        if st.button("🔄 Actualizar datos"):
            invalidar()
            st.rerun()
        if st.button("Cerrar sesión"):
            st.session_state.clear()
            st.rerun()

    hero(usuario, es_admin, hoy_ar())
    if not es_admin and not pagina.startswith(("📚", "📜")):
        banner_pendientes(ctx, usuario, hoy_ar())
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
