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
    exacto, por_tarea = {}, {}
    if df_cat is not None and not df_cat.empty:
        for r in df_cat.itertuples():
            d, t, s, tipo = (str(r.Departamento).strip(), str(r.Tarea).strip(),
                              str(r.Subtarea).strip(), str(r.Tipo).strip())
            exacto[(d, t, s)] = tipo
            por_tarea.setdefault((d, t), tipo)
    return exacto, por_tarea


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
    exacto, por_tarea = _mapas_tipo(df_cat)
    out["Tipo"] = [exacto.get((d, t, s)) or por_tarea.get((d, t)) or C.TIPO_TRABAJO
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
                        horas_dia: float, feriados: set) -> dict:
    ini, fin = rango_mes(anio, mes)
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
    """Para un día: trabajo, horas 'Disponible' y margen de cada persona."""
    filas = []
    for persona, hd in horas_personas.items():
        rp = reg[(reg["Persona"] == persona) & (reg["Fecha"].dt.date == dia)]
        d = por_dia(rp, hd, feriados, dia, dia).iloc[0]
        margen = max(0.0, float(d["Capacidad"]) - float(d[C.TIPO_TRABAJO]))
        filas.append({"Persona": persona, "Trabajo (hs)": round(float(d[C.TIPO_TRABAJO]), 1),
                      "Disponible cargado (hs)": round(float(d[C.TIPO_DISPONIBLE]), 1),
                      "Margen (hs)": round(margen, 1)})
    out = pd.DataFrame(filas)
    return out.sort_values("Margen (hs)", ascending=False).reset_index(drop=True) if not out.empty else out


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
                          feriados: set, departamentos: list[str]):
    """Devuelve (horas, resumen):
    - horas: departamento x semana, horas de TRABAJO.
    - resumen: por semana, capacidad del equipo, trabajo total, horas libres y ocupación %.
    Libre = capacidad del equipo - trabajo total (lo que el equipo podría absorber sin extras)."""
    horas_d, cap_d = matriz_departamentos(reg, anio, mes, horas_personas, feriados, departamentos)
    semanas = semanas_del_mes(anio, mes)
    etiquetas = [s[0] for s in semanas]
    h = pd.DataFrame(0.0, index=horas_d.index, columns=etiquetas)
    capacidad = pd.Series(0.0, index=etiquetas)
    for etiqueta, a, b in semanas:
        cols = [d for d in horas_d.columns if a <= d <= b]
        h[etiqueta] = horas_d[cols].sum(axis=1)
        capacidad[etiqueta] = float(cap_d[cols].sum())
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
