"""Lógica de negocio pura: no usa Streamlit ni Google Sheets, así se puede probar sola.

Reglas:
- Capacidad de un día hábil = HorasDia - horas de ausencia de ese día.
- Disponible y "Gestión y mejoras" (Tipo Disponible) = tiempo libre.
- Horas extra de un día = max(0, trabajo - capacidad del día). En un día no hábil la
  capacidad es 0, así que todo lo trabajado cuenta como extra.
"""
from __future__ import annotations

import calendar
from datetime import date, timedelta

import numpy as np
import pandas as pd

import config as C

# ----------------------------------------------------------------------------
# Fechas
# ----------------------------------------------------------------------------

def es_habil(d: date, feriados: set) -> bool:
    return d.weekday() < 5 and d not in feriados


def rango_mes(anio: int, mes: int) -> tuple[date, date]:
    return date(anio, mes, 1), date(anio, mes, calendar.monthrange(anio, mes)[1])


def dias_habiles(desde: date, hasta: date, feriados: set) -> list[date]:
    out, d = [], desde
    while d <= hasta:
        if es_habil(d, feriados):
            out.append(d)
        d += timedelta(days=1)
    return out


def _parse_fechas(serie: pd.Series) -> pd.Series:
    """ISO (AAAA-MM-DD) primero; si alguien editó a mano en el Sheet, acepta dd/mm/aaaa."""
    s = serie.astype(str).str.strip()
    out = pd.to_datetime(s, format="%Y-%m-%d", errors="coerce")
    falta = out.isna() & s.ne("")
    if falta.any():
        out[falta] = pd.to_datetime(s[falta], dayfirst=True, errors="coerce")
    return out


def feriados_set(df_fer: pd.DataFrame) -> set:
    if df_fer is None or df_fer.empty or "Fecha" not in df_fer.columns:
        return set()
    f = _parse_fechas(df_fer["Fecha"]).dropna()
    return {x.date() for x in f}


def anios_sin_feriados(feriados: set, anios: list[int]) -> list[int]:
    con = {d.year for d in feriados}
    return [a for a in anios if a not in con]

# ----------------------------------------------------------------------------
# Catálogo y personas
# ----------------------------------------------------------------------------

def catalogo_activo(df_cat: pd.DataFrame) -> pd.DataFrame:
    if df_cat is None or df_cat.empty:
        return pd.DataFrame(columns=C.COLS_CATALOGO)
    out = df_cat.copy()
    for c in ("Departamento", "Tarea", "Subtarea", "Tipo", "Activo"):
        out[c] = out[c].fillna("").astype(str).str.strip()
    out["Orden"] = pd.to_numeric(out["Orden"], errors="coerce").fillna(10**6)
    out = out[out["Activo"].str.upper().isin(["SI", "SÍ", "TRUE", "1", ""])]
    return out.sort_values("Orden", kind="stable").reset_index(drop=True)


def personas_df(df_per: pd.DataFrame) -> pd.DataFrame:
    if df_per is None or df_per.empty:
        return pd.DataFrame(columns=["Nombre", "Rol", "Activo", "HorasDia"])
    out = df_per.copy()
    out["Nombre"] = out["Nombre"].fillna("").astype(str).str.strip()
    out["Rol"] = out["Rol"].fillna("").astype(str).str.strip()
    out["Activo"] = out["Activo"].fillna("").astype(str).str.strip().str.upper().isin(["SI", "SÍ", "TRUE", "1"])
    out["HorasDia"] = pd.to_numeric(out["HorasDia"].astype(str).str.replace(",", "."),
                                    errors="coerce").fillna(C.HORAS_DIA_DEFAULT)
    return out[out["Nombre"] != ""].reset_index(drop=True)


def horas_por_persona(personas: pd.DataFrame) -> dict:
    """Operarios activos y sus horas por día (el Admin no cuenta como capacidad)."""
    op = personas[(personas["Rol"] == C.ROL_OPERARIO) & personas["Activo"]]
    return {r.Nombre: float(r.HorasDia) for r in op.itertuples()}

# ----------------------------------------------------------------------------
# Registros
# ----------------------------------------------------------------------------

def _mapas_tipo(df_cat: pd.DataFrame):
    exacto, por_tarea, solo_tarea = {}, {}, {}
    if df_cat is not None and not df_cat.empty:
        for r in df_cat.itertuples():
            d, t, s, tipo = (str(r.Departamento).strip(), str(r.Tarea).strip(),
                              str(r.Subtarea).strip(), str(r.Tipo).strip())
            exacto[(d, t, s)] = tipo
            por_tarea.setdefault((d, t), tipo)
            solo_tarea.setdefault(t, tipo)
    return exacto, por_tarea, solo_tarea


def preparar_registros(df_reg: pd.DataFrame, df_cat: pd.DataFrame) -> pd.DataFrame:
    """Limpia los tipos (el Sheet devuelve todo como texto) y agrega Horas y Tipo."""
    cols = C.COLS_REGISTROS
    base = df_reg.copy() if df_reg is not None else pd.DataFrame(columns=cols)
    for c in cols:
        if c not in base.columns:
            base[c] = ""
    out = base[cols].copy()
    out["Fecha"] = _parse_fechas(out["Fecha"])
    out = out[out["Fecha"].notna()].copy()
    out["Minutos"] = pd.to_numeric(out["Minutos"].astype(str).str.replace(",", "."),
                                   errors="coerce").fillna(0.0)
    out["Horas"] = out["Minutos"] / 60.0
    for c in ("ID", "Persona", "Departamento", "Tarea", "Subtarea", "Nota", "Registrado"):
        out[c] = out[c].fillna("").astype(str).str.strip()
    exacto, por_tarea, solo_tarea = _mapas_tipo(df_cat)
    # Si no hay coincidencia exacta se busca por departamento+tarea y por último solo por tarea
    # (así «Disponible» y «Ausencias» funcionan aunque se carguen sin departamento).
    out["Tipo"] = [exacto.get((d, t, s)) or por_tarea.get((d, t)) or solo_tarea.get(t) or C.TIPO_TRABAJO
                   for d, t, s in zip(out["Departamento"], out["Tarea"], out["Subtarea"])]
    return out.reset_index(drop=True)


def del_mes(reg: pd.DataFrame, anio: int, mes: int) -> pd.DataFrame:
    return reg[(reg["Fecha"].dt.year == anio) & (reg["Fecha"].dt.month == mes)]

# ----------------------------------------------------------------------------
# Capacidad, extras y resumen por persona
# ----------------------------------------------------------------------------

def por_dia(reg_p: pd.DataFrame, horas_dia: float, feriados: set,
            desde: date, hasta: date) -> pd.DataFrame:
    """Una fila por día entre desde y hasta, con Trabajo, Disponible, Ausencia,
    Capacidad y Extra (horas)."""
    idx = pd.date_range(desde, hasta, freq="D")
    tipos = [C.TIPO_TRABAJO, C.TIPO_DISPONIBLE, C.TIPO_AUSENCIA]
    tab = pd.DataFrame(0.0, index=idx, columns=tipos)
    if reg_p is not None and not reg_p.empty:
        g = (reg_p.assign(Dia=reg_p["Fecha"].dt.normalize())
             .groupby(["Dia", "Tipo"])["Horas"].sum())
        for (dia, tipo), h in g.items():
            if dia in tab.index and tipo in tab.columns:
                tab.loc[dia, tipo] = h
    hab = np.array([es_habil(d.date(), feriados) for d in idx])
    cap = np.where(hab, np.maximum(0.0, horas_dia - tab[C.TIPO_AUSENCIA].to_numpy()), 0.0)
    tab["Capacidad"] = cap
    tab["Extra"] = np.maximum(0.0, tab[C.TIPO_TRABAJO].to_numpy() - cap)
    return tab


def resumen_persona_mes(reg: pd.DataFrame, persona: str, anio: int, mes: int,
                        horas_dia: float, feriados: set, hasta: date | None = None) -> dict:
    """'hasta' (opcional) corta el mes en ese día: sirve para medir un mes en curso sin contar los días que faltan."""
    ini, fin = rango_mes(anio, mes)
    if hasta is not None:
        fin = min(fin, hasta)
    rp = reg[(reg["Persona"] == persona)]
    rp = rp[(rp["Fecha"].dt.date >= ini) & (rp["Fecha"].dt.date <= fin)]
    d = por_dia(rp, horas_dia, feriados, ini, fin)
    cap = float(d["Capacidad"].sum())
    trabajo = float(d[C.TIPO_TRABAJO].sum())
    disp = float(d[C.TIPO_DISPONIBLE].sum())
    aus = float(d[C.TIPO_AUSENCIA].sum())
    return {
        "capacidad": round(cap, 2), "trabajo": round(trabajo, 2),
        "disponible": round(disp, 2), "ausencia": round(aus, 2),
        "extra": round(float(d["Extra"].sum()), 2),
        "utilizacion": round(trabajo / cap * 100, 1) if cap > 0 else 0.0,
        "disponibilidad": round(disp / cap * 100, 1) if cap > 0 else 0.0,
        "dias_habiles": len(dias_habiles(ini, fin, feriados)),
        "detalle_dias": d,
    }


def dias_sin_carga(reg: pd.DataFrame, persona: str, anio: int, mes: int,
                   feriados: set, hoy: date) -> list[date]:
    """Días hábiles del mes, hasta AYER, sin ninguna carga. Nunca mira más allá del mes
    ni del día de hoy (el sistema viejo contaba días de meses posteriores)."""
    ini, fin = rango_mes(anio, mes)
    tope = min(fin, hoy - timedelta(days=1))
    if tope < ini:
        return []
    rp = reg[reg["Persona"] == persona]
    cargados = set(rp["Fecha"].dt.date)
    return [d for d in dias_habiles(ini, tope, feriados) if d not in cargados]

# ----------------------------------------------------------------------------
# Generación de registros (un día, rango, vacaciones)
# ----------------------------------------------------------------------------

def repartir_minutos(total: int, n: int, paso: int = C.PASO_MINUTOS) -> list[int]:
    """Reparte 'total' minutos en n días, en bloques de 'paso', sumando exacto."""
    if n <= 0:
        return []
    bloques = int(total) // paso
    base, resto = divmod(bloques, n)
    return [(base + (1 if i < resto else 0)) * paso for i in range(n)]


def fechas_de_carga(desde: date, hasta: date, feriados: set, solo_habiles: bool) -> list[date]:
    """Un solo día se acepta aunque sea fin de semana o feriado (queda como extra).
    En un rango, o con solo_habiles, se saltean fines de semana y feriados."""
    if desde == hasta and not solo_habiles:
        return [desde]
    return dias_habiles(desde, hasta, feriados)


def generar_registros(persona: str, depto: str, tarea: str, subtarea: str, nota: str,
                      desde: date, hasta: date, feriados: set, *, minutos: int,
                      modo: str = "por_dia", solo_habiles: bool = False) -> list[dict]:
    """modo 'por_dia': 'minutos' por cada día. modo 'total': 'minutos' repartidos en el rango."""
    dias = fechas_de_carga(desde, hasta, feriados, solo_habiles)
    if not dias:
        return []
    if modo == "total":
        mins = repartir_minutos(minutos, len(dias))
    else:
        mins = [int(minutos)] * len(dias)
    return [{"Fecha": d.isoformat(), "Persona": persona, "Departamento": depto,
             "Tarea": tarea, "Subtarea": subtarea, "Minutos": int(m), "Nota": nota}
            for d, m in zip(dias, mins) if m > 0]

# ----------------------------------------------------------------------------
# Saturación por departamento (panel de Admin)
# ----------------------------------------------------------------------------

def matriz_departamentos(reg: pd.DataFrame, anio: int, mes: int, horas_personas: dict,
                         feriados: set, departamentos: list[str]):
    """Devuelve (horas, capacidad_equipo):
    - horas: DataFrame departamento x día con las horas de TRABAJO.
    - capacidad_equipo: Serie por día con la capacidad conjunta del equipo."""
    ini, fin = rango_mes(anio, mes)
    dias = [d.date() for d in pd.date_range(ini, fin, freq="D")]
    horas = pd.DataFrame(0.0, index=departamentos, columns=dias)
    rm = reg[(reg["Tipo"] == C.TIPO_TRABAJO)]
    rm = rm[(rm["Fecha"].dt.date >= ini) & (rm["Fecha"].dt.date <= fin)]
    if not rm.empty:
        g = rm.groupby([rm["Departamento"], rm["Fecha"].dt.date])["Horas"].sum()
        for (dep, dia), h in g.items():
            if dep not in horas.index:
                horas.loc[dep] = 0.0
            horas.loc[dep, dia] = h
    cap = pd.Series(0.0, index=dias)
    for persona, hd in horas_personas.items():
        rp = reg[reg["Persona"] == persona]
        d = por_dia(rp, hd, feriados, ini, fin)
        cap = cap + pd.Series(d["Capacidad"].to_numpy(), index=dias)
    return horas, cap


def top_dias_pico(horas: pd.DataFrame, cap_equipo: pd.Series, n: int = 3) -> pd.DataFrame:
    """Los n días de mayor carga de cada departamento."""
    filas = []
    for dep, serie in horas.iterrows():
        for dia, h in serie.nlargest(n).items():
            if h > 0:
                cap = float(cap_equipo.get(dia, 0.0))
                filas.append({"Departamento": dep, "Día": dia, "Horas": round(float(h), 1),
                              "% del equipo": round(h / cap * 100, 0) if cap > 0 else None})
    return pd.DataFrame(filas, columns=["Departamento", "Día", "Horas", "% del equipo"])


def quien_puede_ayudar(reg: pd.DataFrame, dia: date, horas_personas: dict,
                       feriados: set) -> pd.DataFrame:
    """Para un día: trabajo, horas 'Disponible' y margen de cada persona.
    Si alguien no cargó nada ese día, NO se supone que está libre: figura 'sin carga ese día'."""
    filas = []
    for persona, hd in horas_personas.items():
        rp = reg[(reg["Persona"] == persona) & (reg["Fecha"].dt.date == dia)]
        d = por_dia(rp, hd, feriados, dia, dia).iloc[0]
        if rp.empty:
            estado = "día no hábil" if d["Capacidad"] == 0 else "sin carga ese día"
            margen = float("nan")
        else:
            estado = "con carga"
            margen = round(max(0.0, float(d["Capacidad"]) - float(d[C.TIPO_TRABAJO])), 1)
        filas.append({"Persona": persona, "Estado": estado,
                      "Trabajo (hs)": round(float(d[C.TIPO_TRABAJO]), 1),
                      "Disponible cargado (hs)": round(float(d[C.TIPO_DISPONIBLE]), 1),
                      "Margen (hs)": margen})
    out = pd.DataFrame(filas)
    if out.empty:
        return out
    return out.sort_values("Margen (hs)", ascending=False, na_position="last").reset_index(drop=True)


# ----------------------------------------------------------------------------
# Etapa 2: semanas, objetivo del mes, tendencia, desvío
# ----------------------------------------------------------------------------

def meses_atras(anio: int, mes: int, n: int) -> list[tuple[int, int]]:
    """Los últimos n meses (del más viejo al actual), incluyendo el indicado."""
    out, a, m = [], anio, mes
    for _ in range(n):
        out.append((a, m))
        m -= 1
        if m == 0:
            a, m = a - 1, 12
    return out[::-1]


def semanas_del_mes(anio: int, mes: int) -> list[tuple[str, date, date]]:
    """Semanas de lunes a domingo recortadas al mes. Ej.: ('Sem 1 (01/10-04/10)', ini, fin)."""
    ini, fin = rango_mes(anio, mes)
    out, d, n = [], ini, 1
    while d <= fin:
        fin_sem = min(fin, d + timedelta(days=6 - d.weekday()))
        out.append((f"Sem {n} ({d:%d/%m}-{fin_sem:%d/%m})", d, fin_sem))
        d, n = fin_sem + timedelta(days=1), n + 1
    return out


def semanal_departamentos(reg: pd.DataFrame, anio: int, mes: int, horas_personas: dict,
                          feriados: set, departamentos: list[str], solo_con_datos: bool = True):
    """Devuelve (horas, resumen):
    - horas: departamento x semana, horas de TRABAJO.
    - resumen: por semana, capacidad del equipo, trabajo total, horas libres y ocupación %.
    Libre = capacidad del equipo - trabajo total (lo que el equipo podría absorber sin extras).
    Con solo_con_datos (por defecto) la capacidad cuenta únicamente los días en que alguien cargó horas:
    así las semanas que todavía no pasaron no aparecen como 'libres'."""
    horas_d, cap_d = matriz_departamentos(reg, anio, mes, horas_personas, feriados, departamentos)
    semanas = semanas_del_mes(anio, mes)
    con_datos = dias_con_datos(reg, anio, mes)
    etiquetas = [s[0] for s in semanas]
    h = pd.DataFrame(0.0, index=horas_d.index, columns=etiquetas)
    capacidad = pd.Series(0.0, index=etiquetas)
    for etiqueta, a, b in semanas:
        cols = [d for d in horas_d.columns if a <= d <= b]
        h[etiqueta] = horas_d[cols].sum(axis=1)
        cols_cap = [d for d in cols if d in con_datos] if solo_con_datos else cols
        capacidad[etiqueta] = float(cap_d[cols_cap].sum())
    total = h.sum(axis=0)
    libre = (capacidad - total).clip(lower=0)
    ocup = (total / capacidad * 100).where(capacidad > 0)
    resumen = pd.DataFrame({"Capacidad del equipo (hs)": capacidad.round(1),
                            "Trabajo total (hs)": total.round(1),
                            "Libre (hs)": libre.round(1),
                            "Ocupación %": ocup.round(0)})
    return h, resumen


def objetivo_mes(reg: pd.DataFrame, persona: str, anio: int, mes: int,
                 horas_dia: float, feriados: set) -> tuple[float, float]:
    """(horas que hay que cargar en el mes, horas cargadas hasta ahora).
    Objetivo = días hábiles x horas por día. Cargadas incluye Disponible y Ausencias."""
    ini, fin = rango_mes(anio, mes)
    objetivo = len(dias_habiles(ini, fin, feriados)) * horas_dia
    rp = reg[reg["Persona"] == persona]
    rp = rp[(rp["Fecha"].dt.date >= ini) & (rp["Fecha"].dt.date <= fin)]
    return float(objetivo), float(rp["Horas"].sum())


def tendencia_departamentos(reg: pd.DataFrame, persona: str, anio: int, mes: int, n: int = 6):
    """(DataFrame Mes/Departamento/Horas, lista de meses en orden). Solo horas de trabajo."""
    base = reg[(reg["Persona"] == persona) & (reg["Tipo"] == C.TIPO_TRABAJO)]
    meses = meses_atras(anio, mes, n)
    etiquetas = [f"{C.MESES_ES[m][:3]} {a}" for a, m in meses]
    filas = []
    for (a, m), etiqueta in zip(meses, etiquetas):
        g = del_mes(base, a, m).groupby("Departamento")["Horas"].sum()
        filas += [{"Mes": etiqueta, "Departamento": dep, "Horas": round(float(h), 1)} for dep, h in g.items()]
    return pd.DataFrame(filas, columns=["Mes", "Departamento", "Horas"]), etiquetas


def distribucion_semanal(reg: pd.DataFrame, persona: str, anio: int, mes: int):
    """(DataFrame Semana/Departamento/Horas, lista de semanas en orden). Solo trabajo."""
    base = reg[(reg["Persona"] == persona) & (reg["Tipo"] == C.TIPO_TRABAJO)]
    semanas = semanas_del_mes(anio, mes)
    filas = []
    for etiqueta, a, b in semanas:
        r = base[(base["Fecha"].dt.date >= a) & (base["Fecha"].dt.date <= b)]
        g = r.groupby("Departamento")["Horas"].sum()
        filas += [{"Semana": etiqueta, "Departamento": dep, "Horas": round(float(h), 1)} for dep, h in g.items()]
    return pd.DataFrame(filas, columns=["Semana", "Departamento", "Horas"]), [s[0] for s in semanas]


def desvio_historico(reg: pd.DataFrame, persona: str, anio: int, mes: int,
                     nivel: str = "Departamento", n_prev: int = 2):
    """Horas del mes contra el promedio de los n_prev meses anteriores CON datos.
    nivel: 'Departamento' o 'Tarea'. Devuelve (DataFrame, cantidad de meses usados)."""
    cols = ["Departamento"] if nivel == "Departamento" else ["Departamento", "Tarea"]
    base = reg[(reg["Persona"] == persona) & (reg["Tipo"] == C.TIPO_TRABAJO)]
    actual = del_mes(base, anio, mes).groupby(cols)["Horas"].sum()
    series = []
    for a, m in meses_atras(anio, mes, n_prev + 1)[:-1]:
        s = del_mes(base, a, m).groupby(cols)["Horas"].sum()
        if not s.empty:
            series.append(s)
    out = pd.DataFrame({"Actual (hs)": actual})
    if series:
        prom = pd.concat(series, axis=1).fillna(0).mean(axis=1)
        out = out.join(prom.rename("Promedio previo (hs)"), how="outer")
    else:
        out["Promedio previo (hs)"] = float("nan")
    out = out.fillna({"Actual (hs)": 0.0})
    if series:
        out["Promedio previo (hs)"] = out["Promedio previo (hs)"].fillna(0.0)
    out["Desvío (hs)"] = out["Actual (hs)"] - out["Promedio previo (hs)"]
    out["Desvío %"] = (out["Desvío (hs)"] / out["Promedio previo (hs)"] * 100).where(out["Promedio previo (hs)"] > 0)
    out = out.round(1).reset_index()
    out = out.reindex(out["Desvío (hs)"].abs().sort_values(ascending=False, na_position="last").index)
    return out.reset_index(drop=True), len(series)


# ----------------------------------------------------------------------------
# Calendario de saturación por departamento
# ----------------------------------------------------------------------------
TOTAL_EQUIPO = "TOTAL EQUIPO"


def dias_con_datos(reg: pd.DataFrame, anio: int, mes: int) -> set:
    """Días del mes en que alguien del equipo cargó algo (de cualquier tipo). Un día sin ninguna
    carga es 'sin datos': no se puede saber si estuvo libre o si faltó cargar."""
    ini, fin = rango_mes(anio, mes)
    r = reg[(reg["Fecha"].dt.date >= ini) & (reg["Fecha"].dt.date <= fin)]
    return set(r["Fecha"].dt.date)


def _pct(horas: float, cap: float, pico: float, modo: str, hay_datos: bool) -> float:
    """% para pintar una celda. NaN = sin datos (se muestra en gris).
    modo 'equipo': horas / capacidad del equipo ese día. modo 'pico': horas / mayor día del mismo departamento."""
    if not hay_datos:
        return float("nan")
    if modo == "equipo":
        if cap > 0:
            return min(100.0, horas / cap * 100)
        return 100.0 if horas > 0 else float("nan")   # fin de semana o feriado trabajado: todo es extra
    return min(100.0, horas / pico * 100) if pico > 0 else float("nan")


def calendario_z(horas: pd.DataFrame, cap: pd.Series, con_datos: set, modo: str = "equipo"):
    """(z, texto): departamento x día, con una fila TOTAL EQUIPO arriba. z en % (NaN = sin datos)."""
    total = pd.DataFrame([horas.sum(axis=0)], index=[TOTAL_EQUIPO])
    h = pd.concat([total, horas])
    z = pd.DataFrame(float("nan"), index=h.index, columns=h.columns)
    texto = pd.DataFrame("", index=h.index, columns=h.columns)
    for fila in h.index:
        pico = max([float(h.loc[fila, d]) for d in h.columns if d in con_datos] or [0.0])
        for d in h.columns:
            v = float(h.loc[fila, d])
            z.loc[fila, d] = _pct(v, float(cap.get(d, 0.0)), pico, modo, d in con_datos)
            texto.loc[fila, d] = f"{v:.1f}" if v > 0 else ""
    return z, texto


def grilla_calendario(horas_dep: pd.Series, cap: pd.Series, con_datos: set, feriados: set,
                      anio: int, mes: int, modo: str = "equipo"):
    """Calendario de un departamento: (etiquetas de semana, z, texto), matrices semanas x 7 (lun a dom)."""
    ini, fin = rango_mes(anio, mes)
    semanas = semanas_del_mes(anio, mes)
    pico = max([float(horas_dep.get(d, 0.0)) for d in con_datos if ini <= d <= fin] or [0.0])
    z, texto = [], []
    for _, a, _b in semanas:
        lunes = a - timedelta(days=a.weekday())
        fz, ft = [], []
        for i in range(7):
            d = lunes + timedelta(days=i)
            if d < ini or d > fin:
                fz.append(float("nan"))
                ft.append("")
                continue
            h = float(horas_dep.get(d, 0.0))
            if d in feriados and h == 0:
                fz.append(float("nan"))
                ft.append(f"{d.day}<br>feriado")
                continue
            fz.append(_pct(h, float(cap.get(d, 0.0)), pico, modo, d in con_datos))
            ft.append(f"<b>{d.day}</b>" + (f"<br>{h:.1f} hs" if h > 0 else ""))
        z.append(fz)
        texto.append(ft)
    return [x[0] for x in semanas], z, texto


# ----------------------------------------------------------------------------
# Resumen día por día (¿completó las horas del día?)
# ----------------------------------------------------------------------------
def _hs(x: float) -> str:
    return f"{x:.1f}".replace(".", ",")


def _estado_dia(minutos: int, meta_min: float, habil: bool) -> str:
    """Texto de estado de un día según lo cargado (minutos) y lo que correspondía (meta_min)."""
    if minutos == 0:
        return "❌ sin carga"
    if not habil:
        return "⏱️ día no hábil: cuenta como extra"
    if abs(minutos - meta_min) <= 2:
        return "✅ completo"
    if minutos < meta_min:
        return f"⚠️ faltan {_hs((meta_min - minutos) / 60)} hs"
    return f"⏱️ {_hs((minutos - meta_min) / 60)} hs de más"


def resumen_diario(reg: pd.DataFrame, persona: str, anio: int, mes: int,
                   horas_dia: float, feriados: set, hoy: date) -> pd.DataFrame:
    """Una fila por día con carga, más los días hábiles ya pasados sin carga. Dice si el día quedó
    completo (trabajo + disponible + ausencia = horas del día), si faltan horas o si hay de más.
    El día de hoy no figura como 'sin carga' porque todavía se puede cargar."""
    ini, fin = rango_mes(anio, mes)
    rp = reg[(reg["Persona"] == persona) & (reg["Fecha"].dt.date >= ini) & (reg["Fecha"].dt.date <= fin)]
    d = por_dia(rp, horas_dia, feriados, ini, fin)
    filas = []
    for dia, r in d.iterrows():
        f = dia.date()
        total = float(r[C.TIPO_TRABAJO] + r[C.TIPO_DISPONIBLE] + r[C.TIPO_AUSENCIA])
        hab = es_habil(f, feriados)
        if total <= 0 and (not hab or f >= hoy):
            continue
        estado = _estado_dia(round(total * 60), horas_dia * 60 if hab else 0.0, hab)
        filas.append({"Fecha": f, "Trabajo (hs)": round(float(r[C.TIPO_TRABAJO]), 1),
                      "Disponible (hs)": round(float(r[C.TIPO_DISPONIBLE]), 1),
                      "Ausencia (hs)": round(float(r[C.TIPO_AUSENCIA]), 1),
                      "Total (hs)": round(total, 1), "Estado": estado})
    return pd.DataFrame(filas, columns=["Fecha", "Trabajo (hs)", "Disponible (hs)", "Ausencia (hs)",
                                        "Total (hs)", "Estado"])


def estado_dia(reg: pd.DataFrame, dia: date, horas_personas: dict, feriados: set) -> pd.DataFrame:
    """¿Quién completó las horas de un día? Una fila por persona (para el Admin)."""
    hab = es_habil(dia, feriados)
    filas = []
    for persona, hd in horas_personas.items():
        rp = reg[(reg["Persona"] == persona) & (reg["Fecha"].dt.date == dia)]
        d = por_dia(rp, hd, feriados, dia, dia).iloc[0]
        total = float(d[C.TIPO_TRABAJO] + d[C.TIPO_DISPONIBLE] + d[C.TIPO_AUSENCIA])
        minutos = round(total * 60)
        estado = "— día no hábil" if (not hab and minutos == 0) else _estado_dia(minutos, hd * 60 if hab else 0.0, hab)
        filas.append({"Persona": persona, "Cargadas (hs)": round(total, 1),
                      "Corresponden (hs)": round(hd if hab else 0.0, 1), "Estado": estado})
    return pd.DataFrame(filas, columns=["Persona", "Cargadas (hs)", "Corresponden (hs)", "Estado"])


# ----------------------------------------------------------------------------
# Completar el día y días pendientes
# ----------------------------------------------------------------------------
def registro_disponible(persona: str, dia: date, minutos: int, nota: str = "Completar día") -> dict:
    """Una carga de 'Disponible' (sin departamento) para completar las horas de un día."""
    return {"Fecha": dia.isoformat(), "Persona": persona, "Departamento": C.DEPTO_GENERAL,
            "Tarea": C.TAREA_DISPONIBLE, "Subtarea": "", "Minutos": int(minutos), "Nota": nota}


def dias_pendientes(reg: pd.DataFrame, persona: str, horas_dia: float, feriados: set, hoy: date,
                    dias_atras: int = 31, inicio: date = C.FECHA_INICIO) -> list[tuple]:
    """Días hábiles ya pasados (hasta ayer) en que la persona no completó sus horas:
    [(fecha, minutos que faltan)]. Mira hasta 'dias_atras' días hacia atrás y nunca antes de
    'inicio' (config.FECHA_INICIO), para no reclamar días de antes de que el sistema se usara."""
    desde = max(hoy - timedelta(days=dias_atras), inicio)
    hasta = hoy - timedelta(days=1)
    if hasta < desde:
        return []
    base = reg[reg["Persona"] == persona] if reg is not None and not reg.empty else pd.DataFrame(
        columns=["Fecha", "Tipo", "Horas"])
    d = por_dia(base, horas_dia, feriados, desde, hasta)
    meta = horas_dia * 60
    out = []
    for dia, r in d.iterrows():
        if not es_habil(dia.date(), feriados):
            continue
        minutos = round(float(r[C.TIPO_TRABAJO] + r[C.TIPO_DISPONIBLE] + r[C.TIPO_AUSENCIA]) * 60)
        if minutos < meta - 2:
            out.append((dia.date(), int(round(meta - minutos))))
    return out


# ----------------------------------------------------------------------------
# Carga por departamento: resumen del mes, ventanas pico y "¿podemos absorberlo?"
# ----------------------------------------------------------------------------
def _hd_prom(horas_personas: dict) -> float:
    return sum(horas_personas.values()) / len(horas_personas) if horas_personas else C.HORAS_DIA_DEFAULT


def resumen_departamentos_mes(reg: pd.DataFrame, anio: int, mes: int, horas_personas: dict,
                              feriados: set, departamentos: list[str], hasta: date | None = None) -> pd.DataFrame:
    """Horas de trabajo de cada departamento en el mes, su peso, y la comparación con el mes anterior.
    'Personas equivalentes' = horas / (días hábiles x horas por día): cuántas personas a tiempo
    completo hicieron falta. Con 'hasta' el mes se corta en ese día y el mes anterior se compara en el
    mismo tramo (del día 1 al mismo número de día), para no comparar un mes a medias contra uno completo."""
    ini, fin = rango_mes(anio, mes)
    tope = min(fin, hasta) if hasta is not None else fin
    horas, cap = matriz_departamentos(reg, anio, mes, horas_personas, feriados, departamentos)
    a2, m2 = meses_atras(anio, mes, 2)[0]
    previas, _ = matriz_departamentos(reg, a2, m2, horas_personas, feriados, departamentos)
    cols = [d for d in horas.columns if d <= tope]
    cols_prev = [d for d in previas.columns if (d.day <= tope.day if tope < fin else True)]
    total = horas[cols].sum(axis=1)
    prev = previas[cols_prev].sum(axis=1).reindex(total.index).fillna(0.0)
    dias_hab = len(dias_habiles(ini, tope, feriados))
    cap_mes, trabajo_total = float(cap[cols].sum()), float(total.sum())
    base_persona = dias_hab * _hd_prom(horas_personas)
    out = pd.DataFrame({
        "Departamento": total.index,
        "Horas": total.round(1).to_numpy(),
        "% del trabajo": (total / trabajo_total * 100).round(1).to_numpy() if trabajo_total > 0 else 0.0,
        "% de la capacidad del equipo": (total / cap_mes * 100).round(1).to_numpy() if cap_mes > 0 else 0.0,
        "Personas equivalentes": (total / base_persona).round(2).to_numpy() if base_persona > 0 else 0.0,
        "Mes anterior (hs)": prev.round(1).to_numpy(),
    })
    out["Variación %"] = ((out["Horas"] - out["Mes anterior (hs)"]) / out["Mes anterior (hs)"] * 100
                          ).where(out["Mes anterior (hs)"] > 0).round(0)
    return out.sort_values("Horas", ascending=False, kind="stable").reset_index(drop=True)


def ventanas_pico(reg: pd.DataFrame, anio: int, mes: int, horas_personas: dict, feriados: set,
                  departamentos: list[str], n_dias: int = 4) -> pd.DataFrame:
    """Para cada departamento, la ventana de n días hábiles seguidos con más horas de trabajo.
    Solo se miran ventanas donde alguien cargó todos los días (si no, las horas libres saldrían infladas).
    'Una persona alcanza' = lo que una sola persona puede hacer en esos días (n x horas por día).
    'Libre del equipo' = capacidad de todos - trabajo de todos en esos días."""
    horas, cap = matriz_departamentos(reg, anio, mes, horas_personas, feriados, departamentos)
    con_datos = dias_con_datos(reg, anio, mes)
    habiles = [d for d in horas.columns if es_habil(d, feriados)]
    n = min(n_dias, len(habiles))
    cols = ["Departamento", "Desde", "Hasta", "Horas", "Personas equivalentes", "Libre del equipo (hs)", "Lectura"]
    if n <= 0:
        return pd.DataFrame(columns=cols)
    trabajo_total = horas.sum(axis=0)
    una_persona = n * _hd_prom(horas_personas)
    filas = []
    for dep in horas.index:
        mejor = None
        for i in range(len(habiles) - n + 1):
            v = habiles[i:i + n]
            if not all(d in con_datos for d in v):
                continue
            h = float(horas.loc[dep, v].sum())
            if mejor is None or h > mejor[0]:
                mejor = (h, v)
        if mejor is None or mejor[0] <= 0:
            continue
        h, v = mejor
        libre = max(0.0, float(cap[v].sum()) - float(trabajo_total[v].sum()))
        if h <= una_persona + 0.05:
            lectura = "✅ Una persona alcanza"
        else:
            exceso = h - una_persona
            lectura = (f"🤝 Necesita ayuda: el equipo tenía {_hs(libre)} hs libres" if libre >= exceso
                       else f"⚠️ No alcanza: faltan {_hs(exceso - libre)} hs")
        filas.append({"Departamento": dep, "Desde": v[0], "Hasta": v[-1], "Horas": round(h, 1),
                      "Personas equivalentes": round(h / una_persona, 2),
                      "Libre del equipo (hs)": round(libre, 1), "Lectura": lectura})
    out = pd.DataFrame(filas, columns=cols)
    return out.sort_values("Horas", ascending=False, kind="stable").reset_index(drop=True)


def puede_absorber(reg: pd.DataFrame, anio: int, mes: int, horas_personas: dict, feriados: set,
                   horas_necesarias: float, n_dias: int):
    """¿Podría el equipo absorber 'horas_necesarias' en 'n_dias' días hábiles? Toma como referencia las horas
    libres promedio por día del mes indicado (solo días con datos). Devuelve None si no hay datos."""
    horas, cap = matriz_departamentos(reg, anio, mes, horas_personas, feriados, [])
    con_datos = dias_con_datos(reg, anio, mes)
    dias = [d for d in horas.columns if es_habil(d, feriados) and d in con_datos]
    if not dias or n_dias <= 0:
        return None
    trabajo = horas.sum(axis=0)
    libre_prom = sum(max(0.0, float(cap[d]) - float(trabajo[d])) for d in dias) / len(dias)
    cap_prom = sum(float(cap[d]) for d in dias) / len(dias)
    libre_ventana = libre_prom * n_dias
    return {
        "dias_usados": len(dias), "libre_por_dia": libre_prom, "libre_ventana": libre_ventana,
        "personas_equivalentes": horas_necesarias / (n_dias * _hd_prom(horas_personas)),
        "pct_capacidad": (horas_necesarias / (n_dias * cap_prom) * 100) if cap_prom > 0 else None,
        "alcanza": libre_ventana >= horas_necesarias, "faltan": max(0.0, horas_necesarias - libre_ventana),
    }


# ----------------------------------------------------------------------------
# Informe por departamento (para la reunión con el contador)
# ----------------------------------------------------------------------------
def kpis_equipo(reg: pd.DataFrame, anio: int, mes: int, horas_personas: dict, feriados: set,
                hasta: date | None = None):
    """(tabla por persona, totales del equipo)."""
    filas = []
    for p, hd in horas_personas.items():
        r = resumen_persona_mes(reg, p, anio, mes, hd, feriados, hasta)
        filas.append({"Persona": p, "Capacidad (hs)": r["capacidad"], "Trabajo (hs)": r["trabajo"],
                      "Utilización %": r["utilizacion"], "Disponible (hs)": r["disponible"],
                      "Ausencias (hs)": r["ausencia"], "Extra (hs)": r["extra"]})
    tabla = pd.DataFrame(filas, columns=["Persona", "Capacidad (hs)", "Trabajo (hs)", "Utilización %",
                                         "Disponible (hs)", "Ausencias (hs)", "Extra (hs)"])
    tot = {k: round(float(tabla[c].sum()), 1) for k, c in
           [("capacidad", "Capacidad (hs)"), ("trabajo", "Trabajo (hs)"), ("disponible", "Disponible (hs)"),
            ("ausencia", "Ausencias (hs)"), ("extra", "Extra (hs)")]}
    tot["utilizacion"] = round(tot["trabajo"] / tot["capacidad"] * 100, 1) if tot["capacidad"] > 0 else 0.0
    tot["disponibilidad"] = round(tot["disponible"] / tot["capacidad"] * 100, 1) if tot["capacidad"] > 0 else 0.0
    return tabla, tot


def ocupacion_diaria_equipo(reg: pd.DataFrame, anio: int, mes: int, horas_personas: dict,
                            feriados: set, departamentos: list[str]) -> pd.DataFrame:
    """Un día por fila (solo días con carga): capacidad y trabajo del equipo, ocupación %, extra y estado.
    Saturado = ocupación >= UMBRAL_SATURADO o 1 hora extra o más del equipo; Alto = >= UMBRAL_ALTO."""
    horas, cap = matriz_departamentos(reg, anio, mes, horas_personas, feriados, departamentos)
    con_datos = dias_con_datos(reg, anio, mes)
    ini, fin = rango_mes(anio, mes)
    extra = {d: 0.0 for d in horas.columns}
    for persona, hd in horas_personas.items():
        det = por_dia(reg[reg["Persona"] == persona], hd, feriados, ini, fin)
        for dia, v in det["Extra"].items():
            extra[dia.date()] += float(v)
    filas = []
    for dia in horas.columns:
        trabajo = float(horas[dia].sum())
        if dia not in con_datos or (not es_habil(dia, feriados) and trabajo == 0):
            continue
        c = float(cap[dia])
        ocup = trabajo / c * 100 if c > 0 else (100.0 if trabajo > 0 else 0.0)
        estado = ("Saturado" if (ocup >= C.UMBRAL_SATURADO or extra[dia] >= 1.0)
                  else "Alto" if ocup >= C.UMBRAL_ALTO else "Con margen")
        filas.append({"Fecha": dia, "Capacidad (hs)": round(c, 1), "Trabajo (hs)": round(trabajo, 1),
                      "Libre (hs)": round(max(0.0, c - trabajo), 1), "Ocupación %": round(ocup, 0),
                      "Extra (hs)": round(extra[dia], 1),
                      "Departamento principal": str(horas[dia].idxmax()) if trabajo > 0 else "",
                      "Estado": estado})
    return pd.DataFrame(filas, columns=["Fecha", "Capacidad (hs)", "Trabajo (hs)", "Libre (hs)", "Ocupación %",
                                        "Extra (hs)", "Departamento principal", "Estado"])


def cobertura_carga(reg: pd.DataFrame, anio: int, mes: int, horas_personas: dict, feriados: set,
                    hoy: date) -> tuple[int, int]:
    """(días-persona completos, días-persona hábiles) hasta ayer. Mide qué tan confiable es el informe."""
    ini, fin = rango_mes(anio, mes)
    tope = min(fin, hoy - timedelta(days=1))
    if tope < ini:
        return 0, 0
    completos = total = 0
    for p, hd in horas_personas.items():
        d = por_dia(reg[reg["Persona"] == p], hd, feriados, ini, tope)
        for dia, r in d.iterrows():
            if not es_habil(dia.date(), feriados):
                continue
            total += 1
            if round(float(r[C.TIPO_TRABAJO] + r[C.TIPO_DISPONIBLE] + r[C.TIPO_AUSENCIA]) * 60) >= hd * 60 - 2:
                completos += 1
    return completos, total


def tiempo_libre_mes(reg: pd.DataFrame, anio: int, mes: int) -> dict:
    """Horas disponibles del mes: 'Disponible' puro y 'Gestión y mejoras' por departamento."""
    rm = del_mes(reg[reg["Tipo"] == C.TIPO_DISPONIBLE], anio, mes)
    puro = float(rm[rm["Tarea"] == C.TAREA_DISPONIBLE]["Horas"].sum())
    gest = rm[rm["Tarea"] == C.TAREA_GESTION].groupby("Departamento")["Horas"].sum()
    gest = gest.round(1).reset_index().sort_values("Horas", ascending=False)
    return {"disponible": round(puro, 1), "gestion": gest, "gestion_total": round(float(gest["Horas"].sum()), 1)}


def informe_departamentos(reg: pd.DataFrame, anio: int, mes: int, horas_personas: dict, feriados: set,
                          departamentos: list[str], hoy: date, n_dias: int = 4) -> dict:
    """Todo lo que lleva el informe para el contador, más 'hallazgos' ya redactados."""
    ini_m, fin_m = rango_mes(anio, mes)
    con_datos = dias_con_datos(reg, anio, mes)
    ultimo = max(con_datos) if con_datos else None          # los números llegan hasta el último día con datos
    parcial = ultimo is not None and ultimo < fin_m
    personas, tot = kpis_equipo(reg, anio, mes, horas_personas, feriados, ultimo)
    res = resumen_departamentos_mes(reg, anio, mes, horas_personas, feriados, departamentos, ultimo)
    sem_h, sem_res = semanal_departamentos(reg, anio, mes, horas_personas, feriados, departamentos)
    dias = ocupacion_diaria_equipo(reg, anio, mes, horas_personas, feriados, departamentos)
    ventanas = ventanas_pico(reg, anio, mes, horas_personas, feriados, departamentos, n_dias)
    horas_d, cap_d = matriz_departamentos(reg, anio, mes, horas_personas, feriados, departamentos)
    picos = top_dias_pico(horas_d, cap_d, 3)
    libre = tiempo_libre_mes(reg, anio, mes)
    completos, total_dp = cobertura_carga(reg, anio, mes, horas_personas, feriados, hoy)

    h = []
    if parcial:
        h.append(f"El mes está en curso: los números llegan hasta el {ultimo:%d/%m}. La comparación con el mes anterior "
                 f"se hace en el mismo tramo (días 1 a {ultimo.day}).")
    if total_dp > 0 and completos / total_dp < 0.9:
        h.append(f"Atención con los datos: solo {completos} de {total_dp} días-persona hábiles quedaron completos "
                 f"({completos / total_dp * 100:.0f}%). Las horas pueden estar subestimadas.")
    activos = res[res["Horas"] > 0]
    if not activos.empty:
        t = activos.iloc[0]
        h.append(f"El departamento con más horas fue {t['Departamento']}: {_hs(t['Horas'])} hs "
                 f"({_hs(t['% del trabajo'])}% del trabajo del equipo).")
        con_base = activos[(activos["Mes anterior (hs)"] >= 5) & activos["Variación %"].notna()]
        if not con_base.empty:
            sube = con_base.loc[con_base["Variación %"].idxmax()]
            baja = con_base.loc[con_base["Variación %"].idxmin()]
            if sube["Variación %"] >= 20:
                h.append(f"Mayor suba contra el mes anterior: {sube['Departamento']} (+{sube['Variación %']:.0f}%, de "
                         f"{_hs(sube['Mes anterior (hs)'])} a {_hs(sube['Horas'])} hs).")
            if baja["Variación %"] <= -20:
                h.append(f"Mayor baja contra el mes anterior: {baja['Departamento']} ({baja['Variación %']:.0f}%, de "
                         f"{_hs(baja['Mes anterior (hs)'])} a {_hs(baja['Horas'])} hs).")
    if tot["capacidad"] > 0:
        h.append(f"El equipo trabajó {_hs(tot['trabajo'])} hs de {_hs(tot['capacidad'])} de capacidad "
                 f"({tot['utilizacion']:.0f}%) y tuvo {_hs(tot['disponible'])} hs disponibles ({tot['disponibilidad']:.0f}%).")
    if not dias.empty:
        d = dias.loc[dias["Ocupación %"].idxmax()]
        h.append(f"El día de mayor carga del equipo fue el {d['Fecha']:%d/%m}: {d['Ocupación %']:.0f}% de la capacidad "
                 f"({_hs(d['Trabajo (hs)'])} de {_hs(d['Capacidad (hs)'])} hs), sobre todo en {d['Departamento principal']}.")
        n_sat = int((dias["Estado"] == "Saturado").sum())
        n_alto = int((dias["Estado"] == "Alto").sum())
        h.append(f"De {len(dias)} {'día' if len(dias) == 1 else 'días'} con carga, {n_sat} "
                 f"{'fue saturado' if n_sat == 1 else 'fueron saturados'} (desde {C.UMBRAL_SATURADO}% de la capacidad "
                 f"o con 1 hora extra o más), {n_alto} de carga alta (desde {C.UMBRAL_ALTO}%) y "
                 f"{len(dias) - n_sat - n_alto} con margen.")
    if tot["extra"] > 0:
        h.append(f"Hubo {_hs(tot['extra'])} hs extra en el mes.")
    if not ventanas.empty:
        pasan = ventanas[ventanas["Lectura"].str.contains("Necesita ayuda|No alcanza")]
        h.append((f"En una ventana de {n_dias} días hábiles, {', '.join(pasan['Departamento'])} "
                  f"{'superó' if len(pasan) == 1 else 'superaron'} lo que hace una persona (ver tabla de picos).")
                 if not pasan.empty else
                 f"Ningún departamento superó lo que hace una persona en una ventana de {n_dias} días hábiles.")
    return {"anio": anio, "mes": mes, "n_dias": n_dias, "ultimo_dia": ultimo, "parcial": parcial, "personas": personas, "totales": tot, "departamentos": res,
            "semanal_horas": sem_h, "semanal": sem_res, "dias": dias, "ventanas": ventanas, "picos": picos,
            "tiempo_libre": libre, "cobertura": (completos, total_dp), "hallazgos": h}
