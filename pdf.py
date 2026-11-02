"""PDFs con ReportLab. Cada función devuelve los bytes del archivo."""
from __future__ import annotations

import re
from io import BytesIO
from xml.sax.saxutils import escape

import pandas as pd
from reportlab.graphics.charts.barcharts import HorizontalBarChart, VerticalBarChart
from reportlab.graphics.shapes import Drawing, Line, Rect, String
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

import config as C

AZUL = colors.HexColor("#1F4E78")
CELESTE = colors.HexColor("#0077B6")
GRIS = colors.HexColor("#F2F5F8")

_EST = {
    "marca": ParagraphStyle("marca", fontName="Helvetica-Bold", fontSize=20, textColor=AZUL, leading=24),
    "titulo": ParagraphStyle("titulo", fontName="Helvetica-Bold", fontSize=14, textColor=CELESTE, leading=18),
    "sub": ParagraphStyle("sub", fontName="Helvetica", fontSize=10, textColor=colors.HexColor("#555555"), leading=13),
    "h": ParagraphStyle("h", fontName="Helvetica-Bold", fontSize=11, textColor=AZUL, spaceBefore=10, spaceAfter=4,
                        keepWithNext=1),
    "n": ParagraphStyle("n", fontName="Helvetica", fontSize=9, leading=12),
    "celda": ParagraphStyle("celda", fontName="Helvetica", fontSize=8, leading=10),
    "celda_b": ParagraphStyle("celda_b", fontName="Helvetica-Bold", fontSize=8, leading=10, textColor=colors.white),
}


def _limpiar(texto) -> str:
    """Saca emojis y símbolos que la tipografía del PDF no puede dibujar."""
    t = str(texto).replace("−", "-").replace("≈", "aprox.").replace("→", "->")
    t = "".join(ch for ch in t if ord(ch) < 256 or ch in "–—•")
    return re.sub(r"\s+", " ", t).strip()


def _p(texto, estilo="n"):
    return Paragraph(escape(_limpiar(texto)), _EST[estilo])


def _num(x, dec=1):
    try:
        v = float(x)
    except (TypeError, ValueError):
        return ""
    return "—" if v != v else f"{v:.{dec}f}".replace(".", ",")


ROJO_SUAVE, AMARILLO_SUAVE, VERDE_SUAVE = (colors.HexColor("#FDE2E2"), colors.HexColor("#FFF1CC"),
                                           colors.HexColor("#DDF3E4"))


def _color_estado(texto: str):
    t = _limpiar(texto)
    if "No alcanza" in t or "Saturado" in t or "Falta capacidad" in t:
        return ROJO_SUAVE
    if "Necesita ayuda" in t or t == "Alto" or "Se cubre con ayuda" in t:
        return AMARILLO_SUAVE
    if "Alcanza con su equipo" in t or "Una persona alcanza" in t or "Con margen" in t or "Sin presión" in t:
        return VERDE_SUAVE
    return None


def _tabla(encabezado: list[str], filas: list[list], anchos: list[float] | None = None,
           colorear: int | None = None):
    """colorear: número de columna cuyo texto (estado/lectura) pinta la celda en rojo, amarillo o verde."""
    datos = [[_p(h, "celda_b") for h in encabezado]]
    datos += [[_p(c, "celda") for c in f] for f in filas]
    t = Table(datos, colWidths=anchos, repeatRows=1)
    extra = []
    if colorear is not None:
        for i, f in enumerate(filas, start=1):
            c = _color_estado(f[colorear])
            if c is not None:
                extra.append(("BACKGROUND", (colorear, i), (colorear, i), c))
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), AZUL),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, GRIS]),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#CCCCCC")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ] + extra))
    return t


def _documento(titulo: str, subtitulo: str):
    buf = BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=1.6 * cm, rightMargin=1.6 * cm,
                            topMargin=1.5 * cm, bottomMargin=1.5 * cm, title=titulo)
    story = [_p("GRUPO PRESSACCO", "marca"), _p(titulo, "titulo"), _p(subtitulo, "sub"), Spacer(1, 10)]
    return buf, doc, story


def _grafico_departamentos(por_depto: pd.DataFrame):
    """Barras horizontales con el color de cada departamento."""
    d = por_depto.sort_values("Horas")
    alto = max(80, 26 * len(d) + 30)
    dib = Drawing(500, alto)
    bc = HorizontalBarChart()
    bc.x, bc.y, bc.width, bc.height = 150, 15, 300, alto - 25
    bc.data = [[round(float(h), 1) for h in d["Horas"]]]
    bc.categoryAxis.categoryNames = [str(x) for x in d["Departamento"]]
    bc.categoryAxis.labels.fontSize = 8
    bc.valueAxis.labels.fontSize = 7
    bc.valueAxis.valueMin = 0
    bc.barLabelFormat = "%.1f"
    bc.barLabels.fontSize = 7
    bc.barLabels.nudge = 12
    bc.bars.strokeColor = None
    for i, dep in enumerate(d["Departamento"]):
        bc.bars[(0, i)].fillColor = colors.HexColor(C.COLORES_DEPTO.get(dep, "#0077B6"))
    dib.add(bc)
    return dib


def pdf_individual(persona: str, periodo: str, resumen: dict, por_depto: pd.DataFrame,
                   detalle: pd.DataFrame, extras: pd.DataFrame | None = None) -> bytes:
    buf, doc, story = _documento(f"Informe mensual — {persona}", periodo)
    kp = [["Capacidad", "Trabajo", "Utilización", "Disponible", "Ausencias", "Horas extra"],
          [f"{_num(resumen['capacidad'])} hs", f"{_num(resumen['trabajo'])} hs", f"{_num(resumen['utilizacion'], 0)} %",
           f"{_num(resumen['disponible'])} hs", f"{_num(resumen['ausencia'])} hs", f"{_num(resumen['extra'])} hs"]]
    t = Table([[ _p(c, "celda_b") for c in kp[0]], [_p(c, "n") for c in kp[1]]], colWidths=[2.9 * cm] * 6)
    t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), CELESTE), ("BOX", (0, 0), (-1, -1), 0.5, CELESTE),
                           ("ALIGN", (0, 0), (-1, -1), "CENTER"), ("BACKGROUND", (0, 1), (-1, 1), GRIS)]))
    story += [t, Spacer(1, 6)]

    if por_depto is not None and not por_depto.empty:
        story += [_p("Horas de trabajo por departamento", "h"), _grafico_departamentos(por_depto)]
    if detalle is not None and not detalle.empty:
        story += [_p("Detalle por tarea y subtarea", "h")]
        filas = [[r.Departamento, r.Tarea, r.Subtarea or "—", _num(r.Horas)] for r in detalle.itertuples()]
        filas.append(["TOTAL", "", "", _num(detalle["Horas"].sum())])
        story.append(_tabla(["Departamento", "Tarea", "Subtarea", "Horas"], filas,
                            [4.2 * cm, 6 * cm, 5.2 * cm, 2 * cm]))
    if extras is not None and not extras.empty:
        story += [_p("Días con horas extra", "h")]
        story.append(_tabla(list(extras.columns), [[_num(v, 1) if isinstance(v, float) else v for v in f]
                                                   for f in extras.itertuples(index=False)]))
    doc.build(story)
    return buf.getvalue()


def pdf_equipo(periodo: str, tabla_equipo: pd.DataFrame, semanal: pd.DataFrame | None = None,
               picos: pd.DataFrame | None = None) -> bytes:
    buf, doc, story = _documento("Informe del equipo", periodo)
    story += [_p("Capacidad y carga por persona", "h")]
    story.append(_tabla(list(tabla_equipo.columns),
                        [[_num(v, 1) if isinstance(v, float) else v for v in f]
                         for f in tabla_equipo.itertuples(index=False)]))
    if semanal is not None and not semanal.empty:
        story += [_p("Carga del equipo por semana", "h")]
        story.append(_tabla(["Semana"] + list(semanal.columns),
                            [[idx] + [_num(v, 0 if "%" in c else 1) for c, v in zip(semanal.columns, f)]
                             for idx, f in zip(semanal.index, semanal.itertuples(index=False))]))
    if picos is not None and not picos.empty:
        story += [_p("Días de mayor carga por departamento", "h")]
        story.append(_tabla(list(picos.columns), [[v if not isinstance(v, float) else _num(v, 1) for v in f]
                                                  for f in picos.itertuples(index=False)]))
    doc.build(story)
    return buf.getvalue()


PROTOCOLO = [
    ("1. Finalidad", "Transformar la carga de trabajo del estudio en datos que permitan decidir: "
                     "dónde hay saturación, quién puede ayudar y cuándo conviene sumar gente."),
    ("2. Objetivos", "Visibilidad de lo que hace cada departamento. Equilibrio entre las personas. "
                     "Transparencia en cómo se usa el tiempo."),
    ("3. Reglas de carga", "Cargar las 6 horas de cada día hábil antes de las 15 hs. Cuando no se está haciendo "
                           "nada, cargar «Disponible». Los minutos se cargan de a 5, con un mínimo de 10."),
    ("4. Qué departamento elegir", "Se elige según para qué es el trabajo, no según quién lo hace. "
                                   "Ej.: reclamar documentación para poder cargarla va en DOCUMENTACIÓN."),
    ("5. Ausencias", "Los días personales y las vacaciones se cargan en «Ausencias». No cuentan como trabajo y "
                     "bajan la capacidad del mes."),
    ("6. Horas extra", "No hay que marcarlas: si en un día se trabaja más que la capacidad de ese día, o se trabaja "
                       "un fin de semana o feriado, la diferencia se cuenta sola como extra."),
    ("7. Tareas no rutinarias", "Los imprevistos y pedidos especiales se cargan en «Tareas no rutinarias». "
                                "En «Otros imprevistos» la nota es obligatoria."),
]


def pdf_protocolo() -> bytes:
    buf, doc, story = _documento("Protocolo de uso", "CRM de capacidad instalada")
    for titulo, texto in PROTOCOLO:
        story += [_p(titulo, "h"), _p(texto)]
    doc.build(story)
    return buf.getvalue()


# ----------------------------------------------------------------------------
# Informe para la reunión con el contador
# ----------------------------------------------------------------------------
def _grafico_ocupacion(dias: pd.DataFrame):
    """Barras: % de la capacidad del equipo usado cada día, con las líneas de carga alta y saturada."""
    d = dias.sort_values("Fecha")
    dib = Drawing(500, 160)
    bc = VerticalBarChart()
    bc.x, bc.y, bc.width, bc.height = 35, 28, 450, 110
    bc.data = [[min(120.0, float(v)) for v in d["Ocupación %"]]]
    bc.categoryAxis.categoryNames = [f"{x:%d/%m}" for x in d["Fecha"]]
    bc.categoryAxis.labels.fontSize = 6
    bc.categoryAxis.labels.angle = 45
    bc.categoryAxis.labels.dy = -8
    bc.valueAxis.valueMin, bc.valueAxis.valueMax, bc.valueAxis.valueStep = 0, 120, 20
    bc.valueAxis.labels.fontSize = 7
    bc.bars.strokeColor = None
    for i, estado in enumerate(d["Estado"]):
        bc.bars[(0, i)].fillColor = {"Saturado": colors.HexColor("#E63946"), "Alto": colors.HexColor("#F4A261")
                                     }.get(estado, colors.HexColor("#52C48F"))
    dib.add(bc)
    for umbral, color in ((C.UMBRAL_ALTO, "#F4A261"), (C.UMBRAL_SATURADO, "#E63946")):
        y = bc.y + bc.height * umbral / 120
        linea = Line(bc.x, y, bc.x + bc.width, y, strokeColor=colors.HexColor(color), strokeWidth=0.8)
        linea.strokeDashArray = [3, 2]
        dib.add(linea)
        dib.add(String(bc.x + bc.width + 2, y - 2, f"{umbral}%", fontSize=6, fillColor=colors.HexColor(color)))
    return dib


def _grafico_tendencia(tend: pd.DataFrame, meses: list[str]):
    """Barras apiladas: horas de cada departamento mes a mes (colores de cada departamento)."""
    tot = tend.groupby("Departamento")["Horas"].sum().sort_values(ascending=False)
    deps = [d for d in tot.index if tot[d] > 0]
    dib = Drawing(500, 200)
    bc = VerticalBarChart()
    bc.x, bc.y, bc.width, bc.height = 40, 62, 440, 125
    bc.data = [[float(tend[(tend["Departamento"] == dep) & (tend["Mes"] == m)]["Horas"].sum()) for m in meses]
               for dep in deps]
    bc.categoryAxis.categoryNames = meses
    bc.categoryAxis.labels.fontSize = 7
    bc.valueAxis.valueMin = 0
    bc.valueAxis.labels.fontSize = 7
    bc.categoryAxis.style = "stacked"
    for i, dep in enumerate(deps):
        bc.bars[i].fillColor = colors.HexColor(C.COLORES_DEPTO.get(dep, "#888888"))
        bc.bars[i].strokeColor = None
    dib.add(bc)
    for i, dep in enumerate(deps):
        x, y = 40 + (i % 3) * 150, 40 - (i // 3) * 11
        dib.add(Rect(x, y, 7, 7, fillColor=colors.HexColor(C.COLORES_DEPTO.get(dep, "#888888")), strokeColor=None))
        dib.add(String(x + 10, y, _limpiar(dep), fontSize=6.5))
    return dib


CORTO = {"ATENCIÓN AL CLIENTE": "Atenc. cliente", "DOCUMENTACIÓN": "Document.", "CONTABILIDAD": "Contab.",
         "RECURSOS HUMANOS": "RR.HH.", "IMPUESTOS": "Impuestos", "GERENCIAL": "Gerencial", "COOPERATIVA": "Coop.",
         "TRÁMITES": "Trámites", "SUELDOS": "Sueldos"}


def _dia(d) -> str:
    return f"{['lun', 'mar', 'mié', 'jue', 'vie', 'sáb', 'dom'][d.weekday()]} {d:%d/%m}"


def pdf_departamentos(d: dict) -> bytes:
    """Informe de carga por departamento. 'd' es lo que devuelve calc.informe_departamentos."""
    periodo = f"{C.MESES_ES[d['mes']]} {d['anio']}"
    t = d["totales"]
    corte = f" · datos hasta el {d['ultimo_dia']:%d/%m}" if d.get("parcial") else ""
    buf, doc, story = _documento("Informe de carga por departamento",
                                 f"{periodo}{corte} · Para la reunión con el contador")

    kp = [["Capacidad", "Trabajo", "Utilización", "Disponible", "Ausencias", "Horas extra"],
          [f"{_num(t['capacidad'])} hs", f"{_num(t['trabajo'])} hs", f"{_num(t['utilizacion'], 0)} %",
           f"{_num(t['disponible'])} hs", f"{_num(t['ausencia'])} hs", f"{_num(t['extra'])} hs"]]
    tk = Table([[_p(c, "celda_b") for c in kp[0]], [_p(c, "n") for c in kp[1]]], colWidths=[2.95 * cm] * 6)
    tk.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), CELESTE), ("BOX", (0, 0), (-1, -1), 0.5, CELESTE),
                            ("ALIGN", (0, 0), (-1, -1), "CENTER"), ("BACKGROUND", (0, 1), (-1, 1), GRIS)]))
    story += [tk, _p("Lo más importante", "h")]
    story += [_p("• " + x) for x in d["hallazgos"]]

    # --- horas por departamento ---
    res = d["departamentos"]
    act = res[(res["Horas"] > 0) | (res["Mes anterior (hs)"] > 0)]
    story += [_p("Horas de trabajo por departamento", "h")]
    if not act[act["Horas"] > 0].empty:
        story.append(_grafico_departamentos(act[act["Horas"] > 0][["Departamento", "Horas"]]))
    filas = [[r["Departamento"], _num(r["Personas del depto."]), _num(r["Horas"]), _num(r["Ocupación del depto. %"], 0) + "%",
              _num(r["% del trabajo"], 0) + "%", _num(r["Mes anterior (hs)"]),
              "—" if pd.isna(r["Variación %"]) else f"{r['Variación %']:+.0f}%"] for _, r in act.iterrows()]
    story.append(_tabla(["Departamento", "Personas", "Horas", "Ocupación del depto.", "% del trabajo",
                         "Mes anterior (hs)", "Variación"], filas,
                        [4.6 * cm, 1.8 * cm, 1.7 * cm, 2.6 * cm, 2.2 * cm, 2.5 * cm, 2.0 * cm]))
    story.append(_p("Personas = cuántas trabajan normalmente en el departamento (hoja Departamentos). Ocupación del "
                    "departamento = horas ÷ (personas × días hábiles × horas por día). Si el mes está en curso, el mes "
                    "anterior se compara en el mismo tramo.", "sub"))

    # --- señal de capacidad ---
    sen = d["senales"]
    story += [_p("¿Hace falta capacidad? Señal por departamento", "h")]
    if sen.empty:
        story.append(_p("Todavía no hay horas de trabajo para evaluar."))
    else:
        filas = [[r["Departamento"], _num(r["Personas"]), _num(r["Horas"]), _num(r["Capacidad del depto. (hs)"]),
                  _num(r["Ocupación %"], 0) + "%", f"{r['Semanas sobre su capacidad']} de {r['Semanas con datos']}",
                  _num(r["Faltan por semana (hs)"]), _num(r["Equivalente en personas"], 2), r["Señal"]]
                 for _, r in sen.iterrows()]
        story.append(_tabla(["Departamento", "Pers.", "Horas", "Capac. (hs)", "Ocup.", "Semanas sobre su capac.",
                             "Faltan por semana", "Pers. equiv.", "Señal"], filas,
                            [3.3 * cm, 1.2 * cm, 1.4 * cm, 1.7 * cm, 1.4 * cm, 2.2 * cm, 1.7 * cm, 1.7 * cm, 3.2 * cm],
                            colorear=8))
        story.append(_p("Cada semana se compara lo trabajado con lo que puede hacer el equipo del departamento. «Faltan por "
                        "semana» son las horas, en promedio, por encima de esa capacidad; «Pers. equiv.» las expresa en semanas "
                        "completas de una persona. Sin presión = no superó su capacidad. Se cubre con ayuda = la superó pero "
                        "al resto del equipo le sobraban horas. Falta capacidad = la superó y al resto no le alcanzaban las "
                        "horas libres: es la señal a mirar mes a mes antes de pensar en sumar a alguien.", "sub"))

    # --- tendencia ---
    story += [_p("Tendencia de los últimos meses", "h")]
    tend, meses = d["tendencia"], d["meses_tendencia"]
    if len(meses) < 2:
        story.append(_p("Todavía hay un solo mes con datos: la tendencia aparece a partir del segundo mes."))
    else:
        story.append(_grafico_tendencia(tend, meses))
        orden = tend.groupby("Departamento")["Horas"].sum().sort_values(ascending=False)
        filas = []
        for dep in [x for x in orden.index if orden[x] > 0]:
            fila = [dep]
            for m in meses:
                r = tend[(tend["Departamento"] == dep) & (tend["Mes"] == m)].iloc[0]
                fila.append(f"{_num(r['Horas'])} ({_num(r['Ocupación %'], 0)}%)")
            filas.append(fila)
        ancho = (17.6 - 4.4) / len(meses)
        story.append(_tabla(["Departamento"] + meses, filas, [4.4 * cm] + [ancho * cm] * len(meses)))
        story.append(_p("Horas de trabajo de cada departamento y, entre paréntesis, su ocupación (horas ÷ capacidad de su "
                        "equipo). El mes en curso se mide hasta el último día cargado.", "sub"))

    # --- días de mayor carga del equipo ---
    dias = d["dias"]
    story += [_p("Días de mayor carga del equipo", "h")]
    if dias.empty:
        story.append(_p("Todavía no hay días con horas cargadas en este mes."))
    else:
        story.append(_grafico_ocupacion(dias))
        story.append(_p(f"Cada barra es el % de la capacidad del equipo usado en el día. Carga alta desde "
                        f"{C.UMBRAL_ALTO}%; saturado desde {C.UMBRAL_SATURADO}% o con 1 hora extra o más del equipo.", "sub"))
        top = dias.sort_values("Ocupación %", ascending=False, kind="stable").head(8)
        filas = [[_dia(r["Fecha"]), _num(r["Trabajo (hs)"]), _num(r["Capacidad (hs)"]), _num(r["Libre (hs)"]),
                  _num(r["Ocupación %"], 0) + "%", _num(r["Extra (hs)"]), r["Departamento principal"], r["Estado"]]
                 for _, r in top.iterrows()]
        story.append(_tabla(["Día", "Trabajo", "Capacidad", "Libre", "Ocupación", "Extra", "Más cargado en", "Estado"],
                            filas, [2.2 * cm, 1.6 * cm, 1.9 * cm, 1.5 * cm, 2.0 * cm, 1.4 * cm, 4.2 * cm, 2.4 * cm],
                            colorear=7))

    # --- semana por semana ---
    sem, semh = d["semanal"], d["semanal_horas"]
    sem = sem[sem["Capacidad del equipo (hs)"] > 0]          # solo semanas con carga
    semh = semh[list(sem.index)] if not sem.empty else semh.iloc[:, :0]
    story += [_p("Semana por semana", "h")]
    story.append(_tabla(["Semana"] + [c.replace(" (hs)", "").replace("del equipo", "equipo") for c in sem.columns],
                        [[idx] + [_num(v, 0 if "%" in c else 1) for c, v in zip(sem.columns, f)]
                         for idx, f in zip(sem.index, sem.itertuples(index=False))],
                        [4.4 * cm, 3.4 * cm, 3.4 * cm, 3.4 * cm, 3.2 * cm]))
    con_horas = semh[semh.sum(axis=1) > 0]
    if not con_horas.empty:
        story += [Spacer(1, 6), _p("Horas de trabajo de cada departamento por semana", "h")]
        ancho_sem = (17.6 - 4.6) / max(1, len(semh.columns))
        story.append(_tabla(["Departamento"] + [c.split(" (")[0] for c in semh.columns],
                            [[dep] + [_num(v) for v in fila] for dep, fila in zip(con_horas.index,
                                                                                con_horas.itertuples(index=False))],
                            [4.6 * cm] + [ancho_sem * cm] * len(semh.columns)))

    # --- picos ---
    v = d["ventanas"]
    story += [_p(f"Picos: la ventana de {d['n_dias']} días hábiles más cargada de cada departamento", "h")]
    if v.empty:
        story.append(_p(f"Todavía no hay {d['n_dias']} días hábiles seguidos con horas cargadas."))
    else:
        filas = [[r["Departamento"], _num(r["Personas del depto."]), _dia(r["Desde"]), _dia(r["Hasta"]), _num(r["Horas"]),
                  _num(r["Libre del equipo (hs)"]), r["Lectura"]] for _, r in v.iterrows()]
        story.append(_tabla(["Departamento", "Pers.", "Desde", "Hasta", "Horas", "Libre equipo", "Lectura"],
                            filas, [3.8 * cm, 1.2 * cm, 2.0 * cm, 2.0 * cm, 1.4 * cm, 2.0 * cm, 5.2 * cm], colorear=6))
        story.append(_p("«Alcanza con su equipo» compara con lo que pueden hacer las personas del departamento en esos "
                        "días. «Libre equipo» = capacidad de todos − trabajo de todos en esos días.", "sub"))

    # --- horas extra ---
    ex = d["extras"]
    story += [_p("Horas extra por departamento y por persona", "h")]
    if ex.empty:
        story.append(_p("No hubo horas extra en el período."))
    else:
        por_dep = ex.groupby("Departamento")["Extra (hs)"].sum().sort_values(ascending=False)
        por_per = ex.groupby("Persona")["Extra (hs)"].sum().sort_values(ascending=False)
        total_ex = float(por_dep.sum())
        story.append(_tabla(["Departamento", "Extra (hs)", "% del total"],
                            [[k, _num(v), _num(v / total_ex * 100, 0) + "%"] for k, v in por_dep.items()],
                            [7 * cm, 3 * cm, 3 * cm]))
        story.append(Spacer(1, 4))
        story.append(_tabla(["Persona", "Extra (hs)", "% del total"],
                            [[k, _num(v), _num(v / total_ex * 100, 0) + "%"] for k, v in por_per.items()],
                            [7 * cm, 3 * cm, 3 * cm]))
        story.append(_p("Las horas extra de cada día se reparten entre los departamentos en proporción a lo que la persona "
                        "trabajó ese día.", "sub"))

    # --- quién trabaja en qué ---
    cr = d["cruce"]
    deps_cr = [c for c in cr.columns if c != "Total" and float(cr[c].sum()) > 0]
    story += [_p("Quién trabaja en qué (horas del período)", "h")]
    if not deps_cr:
        story.append(_p("Todavía no hay horas de trabajo cargadas."))
    else:
        ancho = (17.6 - 2.6 - 1.4) / len(deps_cr)
        story.append(_tabla(["Persona"] + [CORTO.get(x, x) for x in deps_cr] + ["Total"],
                            [[idx] + [_num(r[c]) for c in deps_cr] + [_num(r["Total"])] for idx, r in cr.iterrows()],
                            [2.6 * cm] + [ancho * cm] * len(deps_cr) + [1.4 * cm]))

    # --- tiempo disponible y por persona ---
    libre = d["tiempo_libre"]
    story += [_p("Tiempo disponible y de mejora", "h")]
    filas = [["Disponible (sin tarea)", _num(libre["disponible"])],
             ["Gestión y mejoras de los departamentos (total)", _num(libre["gestion_total"])]]
    filas += [["   " + r["Departamento"], _num(r["Horas"])] for _, r in libre["gestion"].iterrows()]
    story.append(_tabla(["Concepto", "Horas"], filas, [12 * cm, 3 * cm]))

    story += [_p("Por persona", "h")]
    pe = d["personas"]
    story.append(_tabla(list(pe.columns), [[v if isinstance(v, str) else _num(v, 0 if "%" in c else 1)
                                            for c, v in zip(pe.columns, f)] for f in pe.itertuples(index=False)]))

    cub, tot_dp = d["cobertura"]
    story += [_p("Cómo leer este informe", "h"),
              _p("• Las horas son las que cada persona cargó en el sistema: el informe es tan confiable como la carga diaria."),
              _p(f"• Cobertura de carga del mes: {cub} de {tot_dp} días-persona hábiles quedaron completos."
                 if tot_dp else "• Todavía no hay días hábiles transcurridos para medir la cobertura."),
              _p("• Capacidad = días hábiles × horas por día, menos las ausencias. Disponible = horas libres y de mejora."),
              _p("• Horas extra = trabajo por encima de la capacidad del día, o trabajo en fin de semana o feriado."),
              _p("• Las decisiones conviene tomarlas mirando varios meses seguidos, no uno solo.")]
    doc.build(story)
    return buf.getvalue()
