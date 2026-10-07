"""Capa de datos. Dos implementaciones con la misma interfaz:
- SheetsStore: Google Sheets (producción).
- MemoryStore: en memoria (pruebas).

Reglas de oro:
- Las altas son append (nunca se reescribe una hoja entera).
- Editar y borrar se hace por ID, fila por fila, así no se pisan los datos de otras personas.
- Todo se guarda como texto sin interpretar (RAW), con fechas ISO (AAAA-MM-DD), para que
  la configuración regional del Sheet no cambie día y mes.
- bootstrap() crea las hojas y columnas que falten, sin borrar nada.
"""
from __future__ import annotations

import csv
import glob
import os
import uuid
from datetime import datetime
from zoneinfo import ZoneInfo

import pandas as pd

import config as C
from catalogo_seed import CATALOGO_SEED

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def ahora_iso() -> str:
    return datetime.now(ZoneInfo(C.TZ)).strftime("%Y-%m-%d %H:%M:%S")


def nuevo_id() -> str:
    return uuid.uuid4().hex[:12]


def filas_semilla(hoja: str) -> list[dict]:
    """Datos iniciales para las hojas que arrancan con contenido."""
    if hoja == C.HOJA_CATALOGO:
        return [dict(zip(C.COLS_CATALOGO, f)) for f in CATALOGO_SEED]
    if hoja == C.HOJA_PERSONAS:
        return [dict(zip(C.COLS_PERSONAS, f)) for f in C.PERSONAS_SEED]
    if hoja == C.HOJA_FERIADOS:
        filas = []
        for ruta in sorted(glob.glob(os.path.join(BASE_DIR, "feriados_*.csv"))):
            with open(ruta, encoding="utf-8-sig", newline="") as fh:
                for r in csv.DictReader(fh):
                    filas.append({"Fecha": (r.get("fecha") or "").strip(),
                                  "Motivo": (r.get("motivo") or "").strip()})
        return [f for f in filas if f["Fecha"]]
    return []


def _a_texto(v) -> str:
    return "" if v is None else str(v)


# ----------------------------------------------------------------------------
class MemoryStore:
    """Imita al Sheet (todo se guarda como texto) para poder probar sin conexión."""

    def __init__(self):
        self.hojas: dict[str, dict] = {}   # hoja -> {"cols": [...], "filas": [dict]}

    def bootstrap(self):
        for hoja, cols in C.ESQUEMA.items():
            h = self.hojas.get(hoja)
            if h is None:
                h = self.hojas[hoja] = {"cols": list(cols), "filas": []}
            for c in cols:
                if c not in h["cols"]:
                    h["cols"].append(c)
            if hoja in C.SEMBRAR_SI_VACIA and not h["filas"]:
                self.agregar(hoja, filas_semilla(hoja))

    def leer(self, hoja: str) -> pd.DataFrame:
        h = self.hojas[hoja]
        return pd.DataFrame(h["filas"], columns=h["cols"]).astype(str)

    def agregar(self, hoja: str, filas: list[dict]) -> int:
        h = self.hojas[hoja]
        for f in filas:
            h["filas"].append({c: _a_texto(f.get(c, "")) for c in h["cols"]})
        return len(filas)

    def borrar_por_id(self, hoja: str, ids: list[str]) -> int:
        h = self.hojas[hoja]
        antes = len(h["filas"])
        h["filas"] = [f for f in h["filas"] if f.get("ID") not in set(ids)]
        return antes - len(h["filas"])

    def actualizar_por_id(self, hoja: str, id_: str, cambios: dict) -> bool:
        for f in self.hojas[hoja]["filas"]:
            if f.get("ID") == id_:
                for k, v in cambios.items():
                    if k in f:
                        f[k] = _a_texto(v)
                return True
        return False


# ----------------------------------------------------------------------------
class SheetsStore:
    def __init__(self, spreadsheet):
        self.ss = spreadsheet
        self._ws_cache: dict = {}
        self._header_cache: dict = {}

    # ---- hojas ----
    def _ws(self, hoja: str):
        if hoja not in self._ws_cache:
            self._ws_cache[hoja] = self.ss.worksheet(hoja)
        return self._ws_cache[hoja]

    def _header(self, hoja: str) -> list[str]:
        if hoja not in self._header_cache:
            self._header_cache[hoja] = self._ws(hoja).row_values(1)
        return self._header_cache[hoja]

    def _escribir_encabezado(self, ws, header: list[str]):
        if ws.col_count < len(header):
            ws.resize(cols=len(header))
        ws.update(range_name="A1", values=[header])   # por defecto se guarda como texto (RAW)

    def bootstrap(self):
        """Crea hojas y columnas que falten. No borra ni reordena nada existente."""
        existentes = {w.title: w for w in self.ss.worksheets()}
        for hoja, cols in C.ESQUEMA.items():
            ws = existentes.get(hoja)
            if ws is None:
                ws = self.ss.add_worksheet(title=hoja, rows=2000, cols=max(len(cols), 6))
                self._escribir_encabezado(ws, list(cols))
            else:
                header = [h for h in ws.row_values(1)]
                faltan = [c for c in cols if c not in header]
                if not header:
                    self._escribir_encabezado(ws, list(cols))
                elif faltan:
                    self._escribir_encabezado(ws, header + faltan)
            self._ws_cache[hoja] = ws
            self._header_cache.pop(hoja, None)
            if hoja in C.SEMBRAR_SI_VACIA and len(ws.col_values(1)) <= 1:
                self.agregar(hoja, filas_semilla(hoja))

    # ---- lectura ----
    def leer(self, hoja: str) -> pd.DataFrame:
        valores = self._ws(hoja).get_all_values()
        if not valores:
            return pd.DataFrame()
        header, filas = valores[0], valores[1:]
        n = len(header)
        filas = [(f + [""] * n)[:n] for f in filas if any(str(x).strip() for x in f)]
        df = pd.DataFrame(filas, columns=header)
        return df.loc[:, [c for c in df.columns if str(c).strip()]]

    # ---- escritura ----
    def agregar(self, hoja: str, filas: list[dict]) -> int:
        if not filas:
            return 0
        header = self._header(hoja)
        matriz = [[_a_texto(f.get(c, "")) for c in header] for f in filas]
        self._ws(hoja).append_rows(matriz, value_input_option="RAW")
        return len(filas)

    def _filas_por_id(self, hoja: str, ids: list[str]) -> list[int]:
        header = self._header(hoja)
        col = header.index("ID") + 1
        valores = self._ws(hoja).col_values(col)
        buscados = set(ids)
        return [i + 1 for i, v in enumerate(valores) if i > 0 and v in buscados]

    def borrar_por_id(self, hoja: str, ids: list[str]) -> int:
        filas = self._filas_por_id(hoja, ids)
        if not filas:
            return 0
        ws = self._ws(hoja)
        pedidos = [{"deleteDimension": {"range": {
            "sheetId": ws.id, "dimension": "ROWS", "startIndex": r - 1, "endIndex": r}}}
            for r in sorted(filas, reverse=True)]
        self.ss.batch_update({"requests": pedidos})
        return len(filas)

    def actualizar_por_id(self, hoja: str, id_: str, cambios: dict) -> bool:
        from gspread.utils import rowcol_to_a1
        filas = self._filas_por_id(hoja, [id_])
        if not filas:
            return False
        header = self._header(hoja)
        datos = [{"range": rowcol_to_a1(filas[0], header.index(k) + 1),
                  "values": [[_a_texto(v)]]} for k, v in cambios.items() if k in header]
        if datos:
            self._ws(hoja).batch_update(datos)   # por defecto se guarda como texto (RAW)
        return True
