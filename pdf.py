"""PDFs con ReportLab. Cada función devuelve los bytes del archivo."""
from __future__ import annotations

from io import BytesIO
from xml.sax.saxutils import escape

import pandas as pd
from reportlab.graphics.charts.barcharts import HorizontalBarChart
from reportlab.graphics.shapes import Drawing
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
    "h": ParagraphStyle("h", fontName="Helvetica-Bold", fontSize=11, textColor=AZUL, spaceBefore=10, spaceAfter=4),
    "n": ParagraphStyle("n", fontName="Helvetica", fontSize=9, leading=12),
    "celda": ParagraphStyle("celda", fontName="Helvetica", fontSize=8, leading=10),
    "celda_b": ParagraphStyle("celda_b", fontName="Helvetica-Bold", fontSize=8, leading=10, textColor=colors.white),
}


def _p(texto, estilo="n"):
    return Paragraph(escape(str(texto)), _EST[estilo])


def _num(x, dec=1):
    try:
        return f"{float(x):.{dec}f}".replace(".", ",")
    except (TypeError, ValueError):
        return ""


def _tabla(encabezado: list[str], filas: list[list], anchos: list[float] | None = None):
    datos = [[_p(h, "celda_b") for h in encabezado]]
    datos += [[_p(c, "celda") for c in f] for f in filas]
    t = Table(datos, colWidths=anchos, repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), AZUL),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, GRIS]),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#CCCCCC")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]))
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
