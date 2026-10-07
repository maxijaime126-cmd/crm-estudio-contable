"""CRM Grupo Pressacco v2 — capacidad instalada por departamento.

Departamento -> Tarea -> Subtarea, minutos, ausencias, horas extra automáticas y
calendario de saturación por departamento. La lógica está en calc.py (con pruebas) y los
datos en store.py (crea las hojas que falten).
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
import store as S

st.set_page_config(page_title="CRM Grupo Pressacco", layout="wide", page_icon="🏛️")

st.markdown("""
<style>
    .kpi-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
                gap: 10px; margin-bottom: 14px; }
    .kpi-card { background: linear-gradient(135deg, #0077b6, #00b4d8); border-radius: 12px;
                padding: 14px 12px; color: white; text-align: center; }
    .kpi-card h2 { font-size: 1.7rem; margin: 0; font-weight: 700; white-space: nowrap; }
    .kpi-card p  { font-size: 0.82rem; margin: 2px 0 0 0; opacity: 0.9; }
    .kpi-card.extra { background: linear-gradient(135deg, #e76f51, #f4a261); }
    .alerta-box { border-left: 5px solid #e63946; background: #fff0f0; border-radius: 6px;
                  padding: 10px 15px; margin-bottom: 8px; font-size: 0.9rem; color: #333; }
</style>
""", unsafe_allow_html=True)

DIAS_ES = ["lun", "mar", "mié", "jue", "vie", "sáb", "dom"]


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


def kpis(items: list[tuple]):
    """items: (valor, etiqueta, clase_extra). Una sola pieza de HTML, se adapta al ancho."""
    cards = "".join(
        f'<div class="kpi-card {html.escape(c)}"><h2>{html.escape(str(v))}</h2>'
        f'<p>{html.escape(l)}</p></div>' for v, l, c in items)
    st.markdown(f'<div class="kpi-grid">{cards}</div>', unsafe_allow_html=True)


def alertas(mensajes: list[str]):
    """Un solo bloque (no un elemento por alerta): evita errores de renderizado."""
    if mensajes:
        cuerpo = "".join(f'<div class="alerta-box">{html.escape(m)}</div>' for m in mensajes)
        st.markdown(cuerpo, unsafe_allow_html=True)


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
    st.title("🏛️ CRM Grupo Pressacco")
    activos = personas[personas["Activo"]]["Nombre"].tolist()
    u = st.selectbox("Usuario:", ["Seleccionar..."] + activos)
    pwd = st.text_input("Contraseña:", type="password")
    if st.button("Ingresar") and u != "Seleccionar...":
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

    # --- Departamento > Tarea > Subtarea ---
    c1, c2, c3 = st.columns(3)
    deptos = list(dict.fromkeys(cat["Departamento"]))
    depto = c1.selectbox("Departamento", deptos, key="c_dep")
    sub_d = cat[cat["Departamento"] == depto]
    tarea = c2.selectbox("Tarea", list(dict.fromkeys(sub_d["Tarea"])), key=f"c_tar_{depto}")
    subs = [s for s in sub_d[sub_d["Tarea"] == tarea]["Subtarea"] if s]
    if subs:
        subtarea = c3.selectbox("Subtarea", subs, key=f"c_sub_{depto}_{tarea}")
    else:
        subtarea = ""
        c3.caption("Esta tarea no tiene subtareas.")
    fila_cat = sub_d[(sub_d["Tarea"] == tarea) & (sub_d["Subtarea"] == subtarea)]
    tipo = fila_cat["Tipo"].iloc[0] if not fila_cat.empty else C.TIPO_TRABAJO
    es_ausencia = tipo == C.TIPO_AUSENCIA

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

    seccion_mis_cargas(reg, persona)


def seccion_mis_cargas(reg: pd.DataFrame, persona: str):
    st.divider()
    st.subheader("🗂️ Cargas recientes")
    mis = reg[reg["Persona"] == persona].sort_values(["Fecha", "Registrado"], ascending=False).head(40)
    if mis.empty:
        st.caption("Todavía no hay cargas.")
        return
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
# Resumen y paneles
# ============================================================================
def pantalla_resumen(ctx: dict, usuario: str, es_admin: bool):
    reg = ctx["reg"]
    hoy = hoy_ar()
    st.header("📊 Análisis de capacidad")
    anios = sorted({hoy.year - 1, hoy.year, hoy.year + 1, *reg["Fecha"].dt.year.astype(int).tolist()})
    f1, f2, _ = st.columns([1, 1, 3])
    anio = f1.selectbox("Año", anios, index=anios.index(hoy.year))
    mes = f2.selectbox("Mes", list(range(1, 13)), index=hoy.month - 1, format_func=lambda m: C.MESES_ES[m])

    if es_admin:
        t1, t2, t3 = st.tabs(["👤 Individual", "🌐 Equipo", "🗓️ Calendario por departamento"])
        with t1:
            tab_individual(ctx, None, anio, mes, hoy, es_admin)
        with t2:
            tab_equipo(ctx, anio, mes)
        with t3:
            tab_calendario(ctx, anio, mes, hoy)
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

    kpis([
        (fmt_hs(r["capacidad"]), "Capacidad del mes (hs)", ""),
        (fmt_hs(r["trabajo"]), "Trabajo (hs)", ""),
        (f"{r['utilizacion']:.0f}%".replace(".", ","), "Utilización", ""),
        (fmt_hs(r["disponible"]), "Disponible (hs)", ""),
        (f"{r['disponibilidad']:.0f}%", "Disponibilidad", ""),
        (fmt_hs(r["ausencia"]), "Ausencias (hs)", ""),
        (fmt_hs(r["extra"]), "Horas extra", "extra" if r["extra"] > 0 else ""),
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
    if trabajo.empty:
        st.info("Sin horas de trabajo cargadas en este mes.")
    else:
        por_dep = trabajo.groupby("Departamento", as_index=False)["Horas"].sum().sort_values("Horas")
        fig = px.bar(por_dep, x="Horas", y="Departamento", orientation="h", color="Departamento",
                     color_discrete_map=C.COLORES_DEPTO, text=por_dep["Horas"].round(1),
                     title=f"Horas de trabajo por departamento — {persona}")
        fig.update_layout(showlegend=False, margin=dict(l=0, r=0, t=40, b=0))
        st.plotly_chart(fig)
        dep_sel = st.selectbox("Ver detalle de:", list(por_dep["Departamento"][::-1]), key="ind_dep")
        det = (trabajo[trabajo["Departamento"] == dep_sel]
               .groupby(["Tarea", "Subtarea"], as_index=False)["Horas"].sum().sort_values("Horas", ascending=False))
        det["Horas"] = det["Horas"].round(1)
        det["% del departamento"] = (det["Horas"] / det["Horas"].sum() * 100).round(0)
        st.dataframe(det, hide_index=True)

    libres = rm[rm["Tipo"] != C.TIPO_TRABAJO]
    if not libres.empty:
        with st.expander("Disponible y ausencias del mes"):
            tl = libres.groupby(["Tipo", "Departamento", "Tarea", "Subtarea"], as_index=False)["Horas"].sum()
            tl["Horas"] = tl["Horas"].round(1)
            st.dataframe(tl, hide_index=True)

    if r["extra"] > 0:
        with st.expander("⏱️ Días con horas extra"):
            d = r["detalle_dias"]
            d = d[d["Extra"] > 0].reset_index(names="Fecha")
            st.dataframe(pd.DataFrame({
                "Fecha": d["Fecha"].dt.strftime("%d/%m/%Y"), "Trabajo (hs)": d[C.TIPO_TRABAJO].round(1),
                "Capacidad del día (hs)": d["Capacidad"].round(1), "Extra (hs)": d["Extra"].round(1)}),
                hide_index=True)


def tab_equipo(ctx, anio, mes):
    reg, feriados = ctx["reg"], ctx["feriados"]
    horas_pp = calc.horas_por_persona(ctx["personas"])
    filas = []
    for p, hd in horas_pp.items():
        r = calc.resumen_persona_mes(reg, p, anio, mes, hd, feriados)
        filas.append({"Persona": p, "Capacidad (hs)": r["capacidad"], "Trabajo (hs)": r["trabajo"],
                      "Utilización %": r["utilizacion"], "Disponible (hs)": r["disponible"],
                      "Disponibilidad %": r["disponibilidad"], "Ausencias (hs)": r["ausencia"],
                      "Extra (hs)": r["extra"]})
    if not filas:
        st.info("No hay operarios activos.")
        return
    tabla = pd.DataFrame(filas)
    st.dataframe(tabla, hide_index=True)
    st.caption(f"Equipo: capacidad {fmt_hs(tabla['Capacidad (hs)'].sum())} hs · trabajo "
               f"{fmt_hs(tabla['Trabajo (hs)'].sum())} hs · extra {fmt_hs(tabla['Extra (hs)'].sum())} hs.")


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
        z=z, x=[f"{DIAS_ES[d.weekday()]} {d:%d/%m}" for d in dias], y=list(horas.index), text=texto,
        texttemplate="%{text}", zmin=0, zmax=100, colorbar=dict(title="% del pico"),
        colorscale=[[0, "#d8f3dc"], [0.5, "#ffe8a3"], [1, "#e63946"]]))
    fig.update_yaxes(autorange="reversed")
    fig.update_xaxes(type="category", tickangle=-60)
    fig.update_layout(height=max(320, 42 * len(horas.index) + 140), margin=dict(l=0, r=0, t=10, b=0))
    st.plotly_chart(fig)

    st.subheader("Días de mayor carga por departamento")
    picos = calc.top_dias_pico(horas, cap, 3)
    if not picos.empty:
        picos["Día"] = picos["Día"].map(lambda d: f"{DIAS_ES[d.weekday()]} {d:%d/%m}")
        st.dataframe(picos, hide_index=True)

    st.subheader("¿Quién puede ayudar un día puntual?")
    ini, fin = calc.rango_mes(anio, mes)
    dia = st.date_input("Día", value=min(max(hoy, ini), fin), min_value=ini, max_value=fin,
                        format="DD/MM/YYYY", key="cal_dia")
    st.dataframe(calc.quien_puede_ayudar(reg, dia, horas_pp, feriados), hide_index=True)


# ============================================================================
# Manual
# ============================================================================
def pantalla_manual():
    st.header("📚 Manual")
    st.markdown("""
**Cargar horas:** elegí *Departamento → Tarea → Subtarea*. El departamento es **para qué es el trabajo**,
no quién lo hace (por ejemplo, reclamar facturas va en DOCUMENTACIÓN aunque lo haga Atención al cliente).

**Minutos:** el mínimo es 10 y se carga de a 5.

**Disponible:** cuando no estás haciendo nada, cargalo como *Disponible* para completar el día.

**Ausencias:** *Inasistencia* o *Vacaciones*. Bajan tu capacidad del mes y no cuentan como trabajo.
Para vacaciones elegí desde y hasta: se cargan solo los días hábiles.

**Horas extra:** no hay que marcar nada. Si en un día trabajás más de tu capacidad (o trabajás un fin de
semana o feriado), la diferencia se cuenta sola como extra.

**Errores de carga:** en *Cargas recientes* podés editar los minutos y la nota, o eliminar una carga.
""")


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
        st.markdown(f"**{usuario}**")
        pagina = st.radio("Navegación", ["➕ Cargar horas", "📊 Panel de control", "📚 Manual"])
        if st.button("🔄 Actualizar datos"):
            invalidar()
            st.rerun()
        if st.button("Cerrar sesión"):
            st.session_state.clear()
            st.rerun()

    if pagina.startswith("➕"):
        pantalla_carga(ctx, usuario, es_admin)
    elif pagina.startswith("📊"):
        pantalla_resumen(ctx, usuario, es_admin)
    else:
        pantalla_manual()


if __name__ == "__main__":
    main()
