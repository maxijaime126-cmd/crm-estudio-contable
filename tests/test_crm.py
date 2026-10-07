"""Pruebas: python -m unittest discover -s tests -v   (desde la carpeta crm_v2)"""
import os
import sys
import unittest
from datetime import date

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import calc  # noqa: E402
import config as C  # noqa: E402
import store as S  # noqa: E402
from catalogo_seed import CATALOGO_SEED  # noqa: E402

# Octubre 2026: 1/oct es jueves. Feriado el lunes 12/10.
FER = {date(2026, 10, 12)}
HD = 6.0


def catalogo():
    return pd.DataFrame(CATALOGO_SEED, columns=C.COLS_CATALOGO)


def registros(filas):
    """filas: (fecha_iso, persona, depto, tarea, sub, minutos)"""
    df = pd.DataFrame([{"ID": str(i), "Fecha": f, "Persona": p, "Departamento": d, "Tarea": t,
                        "Subtarea": s, "Minutos": str(m), "Nota": "", "Registrado": ""}
                       for i, (f, p, d, t, s, m) in enumerate(filas)])
    return calc.preparar_registros(df, catalogo())


class TestCatalogo(unittest.TestCase):
    def test_semilla(self):
        df = pd.DataFrame(CATALOGO_SEED, columns=C.COLS_CATALOGO)
        self.assertEqual(df["Departamento"].nunique(), 9)
        self.assertEqual(len(df), 203)
        # Ausencias, Disponible y Gestión en los lugares correctos
        for dep in df["Departamento"].unique():
            sub = df[df["Departamento"] == dep]
            self.assertIn("Ausencias", set(sub["Tarea"]))
            self.assertIn("Disponible", set(sub["Tarea"]))
        self.assertNotIn("Gestión y mejoras del departamento", set(df[df.Departamento == "GERENCIAL"]["Tarea"]))
        self.assertNotIn("Tareas no rutinarias", set(df[df.Departamento == "RECURSOS HUMANOS"]["Tarea"]))
        tipos = df.groupby("Tarea")["Tipo"].first()
        self.assertEqual(tipos["Disponible"], C.TIPO_DISPONIBLE)
        self.assertEqual(tipos["Gestión y mejoras del departamento"], C.TIPO_DISPONIBLE)
        self.assertEqual(tipos["Ausencias"], C.TIPO_AUSENCIA)
        self.assertEqual(tipos["Reuniones"], C.TIPO_TRABAJO)


class TestFechas(unittest.TestCase):
    def test_dias_habiles_y_capacidad(self):
        hab = calc.dias_habiles(date(2026, 10, 1), date(2026, 10, 31), FER)
        self.assertEqual(len(hab), 21)  # 22 días de lunes a viernes menos el feriado del 12
        self.assertNotIn(date(2026, 10, 12), hab)

    def test_parse_acepta_iso_y_ddmm(self):
        s = pd.Series(["2026-10-07", "07/10/2026", "", "basura"])
        r = calc._parse_fechas(s)
        self.assertEqual(r[0], pd.Timestamp(2026, 10, 7))
        self.assertEqual(r[1], pd.Timestamp(2026, 10, 7))
        self.assertTrue(pd.isna(r[2]) and pd.isna(r[3]))

    def test_anios_sin_feriados(self):
        self.assertEqual(calc.anios_sin_feriados(FER, [2026, 2027]), [2027])


class TestResumen(unittest.TestCase):
    def test_capacidad_vacaciones_y_extra(self):
        reg = registros([
            ("2026-10-05", "Natalia", "IMPUESTOS", "IVA mensual", "Liquidación", 480),   # 8 hs lunes
            ("2026-10-06", "Natalia", "IMPUESTOS", "IVA mensual", "Control", 360),       # 6 hs
            ("2026-10-07", "Natalia", "IMPUESTOS", "Disponible", "", 120),               # 2 hs libres
            ("2026-10-08", "Natalia", "IMPUESTOS", "Ausencias", "Vacaciones", 360),      # día completo
            ("2026-10-10", "Natalia", "IMPUESTOS", "IVA mensual", "Control", 180),       # sábado, 3 hs
        ])
        r = calc.resumen_persona_mes(reg, "Natalia", 2026, 10, HD, FER)
        self.assertEqual(r["capacidad"], 21 * 6 - 6)            # baja 6 hs por la vacación
        self.assertEqual(r["trabajo"], 8 + 6 + 3)
        self.assertEqual(r["disponible"], 2)
        self.assertEqual(r["ausencia"], 6)
        self.assertEqual(r["extra"], 2 + 3)                      # 2 del lunes + 3 del sábado
        self.assertAlmostEqual(r["utilizacion"], 17 / 120 * 100, places=1)

    def test_ausencia_parcial_baja_capacidad_del_dia(self):
        reg = registros([
            ("2026-10-05", "Natalia", "IMPUESTOS", "Ausencias", "Inasistencia (día personal)", 180),
            ("2026-10-05", "Natalia", "IMPUESTOS", "IVA mensual", "Control", 360),
        ])
        d = calc.resumen_persona_mes(reg, "Natalia", 2026, 10, HD, FER)["detalle_dias"]
        fila = d.loc[pd.Timestamp(2026, 10, 5)]
        self.assertEqual(fila["Capacidad"], 3)
        self.assertEqual(fila["Extra"], 3)

    def test_feriado_trabajado_es_todo_extra(self):
        reg = registros([("2026-10-12", "Natalia", "IMPUESTOS", "IVA mensual", "Control", 240)])
        r = calc.resumen_persona_mes(reg, "Natalia", 2026, 10, HD, FER)
        self.assertEqual(r["extra"], 4)

    def test_sin_datos(self):
        reg = registros([])
        r = calc.resumen_persona_mes(reg, "Natalia", 2026, 10, HD, FER)
        self.assertEqual((r["trabajo"], r["extra"], r["utilizacion"]), (0, 0, 0))
        self.assertEqual(r["capacidad"], 126)  # igual que el "126 hs" del panel viejo


class TestDiasSinCarga(unittest.TestCase):
    def test_no_cuenta_dias_posteriores_al_mes(self):
        """Bug del sistema viejo: en un mes pasado listaba días de meses posteriores."""
        reg = registros([("2026-08-03", "Natalia", "IMPUESTOS", "IVA mensual", "Control", 60)])
        falta = calc.dias_sin_carga(reg, "Natalia", 2026, 8, set(), date(2026, 10, 7))
        self.assertTrue(all(d.month == 8 for d in falta))
        self.assertNotIn(date(2026, 8, 3), falta)
        self.assertEqual(len(falta), 21 - 1)  # agosto 2026 tiene 21 hábiles

    def test_mes_en_curso_hasta_ayer(self):
        reg = registros([])
        falta = calc.dias_sin_carga(reg, "Natalia", 2026, 10, FER, date(2026, 10, 7))
        self.assertEqual(max(falta), date(2026, 10, 6))
        self.assertEqual(calc.dias_sin_carga(reg, "Natalia", 2026, 11, FER, date(2026, 10, 7)), [])


class TestGenerar(unittest.TestCase):
    def test_repartir_exacto(self):
        m = calc.repartir_minutos(500, 3)
        self.assertEqual(sum(m), 500)
        self.assertTrue(all(x % 5 == 0 for x in m))

    def test_vacaciones_solo_habiles(self):
        filas = calc.generar_registros("Natalia", "IMPUESTOS", "Ausencias", "Vacaciones", "",
                                       date(2026, 10, 9), date(2026, 10, 14), FER,
                                       minutos=360, solo_habiles=True)
        # vie 9, (sáb 10, dom 11, lun 12 feriado), mar 13, mié 14
        self.assertEqual([f["Fecha"] for f in filas], ["2026-10-09", "2026-10-13", "2026-10-14"])

    def test_un_dia_no_habil_se_permite_para_trabajo(self):
        filas = calc.generar_registros("Natalia", "IMPUESTOS", "IVA mensual", "Control", "",
                                       date(2026, 10, 10), date(2026, 10, 10), FER, minutos=60)
        self.assertEqual(len(filas), 1)

    def test_rango_total(self):
        filas = calc.generar_registros("Natalia", "IMPUESTOS", "IVA mensual", "Control", "",
                                       date(2026, 10, 5), date(2026, 10, 9), FER,
                                       minutos=605, modo="total")
        self.assertEqual(sum(f["Minutos"] for f in filas), 605 // 5 * 5)


class TestTiposYSaturacion(unittest.TestCase):
    def test_tipo_desde_catalogo_y_minutos_como_texto(self):
        reg = registros([("2026-10-05", "Natalia", "ATENCIÓN AL CLIENTE", "Disponible", "", 90)])
        self.assertEqual(reg.loc[0, "Tipo"], C.TIPO_DISPONIBLE)
        self.assertEqual(reg.loc[0, "Horas"], 1.5)

    def test_tarea_desconocida_cuenta_como_trabajo(self):
        reg = registros([("2026-10-05", "Natalia", "IMPUESTOS", "Tarea vieja", "", 60)])
        self.assertEqual(reg.loc[0, "Tipo"], C.TIPO_TRABAJO)

    def test_matriz_y_picos(self):
        reg = registros([
            ("2026-10-05", "Natalia", "IMPUESTOS", "IVA mensual", "Control", 360),
            ("2026-10-05", "Athina", "IMPUESTOS", "IVA mensual", "Control", 240),
            ("2026-10-06", "Athina", "IMPUESTOS", "IVA mensual", "Control", 120),
            ("2026-10-06", "Natalia", "DOCUMENTACIÓN", "Carga de documentación", "Sociedades", 300),
            ("2026-10-07", "Natalia", "IMPUESTOS", "Disponible", "", 300),   # no es trabajo
        ])
        hp = {"Natalia": 6.0, "Athina": 6.0}
        horas, cap = calc.matriz_departamentos(reg, 2026, 10, hp, FER, ["IMPUESTOS", "DOCUMENTACIÓN"])
        self.assertEqual(horas.loc["IMPUESTOS", date(2026, 10, 5)], 10)
        self.assertEqual(horas.loc["IMPUESTOS", date(2026, 10, 7)], 0)
        self.assertEqual(cap[date(2026, 10, 5)], 12)
        self.assertEqual(cap[date(2026, 10, 10)], 0)   # sábado
        picos = calc.top_dias_pico(horas, cap, 1)
        imp = picos[picos["Departamento"] == "IMPUESTOS"].iloc[0]
        self.assertEqual((imp["Día"], imp["Horas"]), (date(2026, 10, 5), 10.0))
        self.assertEqual(imp["% del equipo"], round(10 / 12 * 100))

    def test_quien_puede_ayudar(self):
        reg = registros([
            ("2026-10-05", "Natalia", "IMPUESTOS", "IVA mensual", "Control", 420),
            ("2026-10-05", "Athina", "IMPUESTOS", "Disponible", "", 240),
        ])
        r = calc.quien_puede_ayudar(reg, date(2026, 10, 5), {"Natalia": 6.0, "Athina": 6.0}, FER)
        self.assertEqual(r.iloc[0]["Persona"], "Athina")
        self.assertEqual(r.iloc[0]["Margen (hs)"], 6.0)
        self.assertEqual(r[r.Persona == "Natalia"].iloc[0]["Margen (hs)"], 0.0)


class TestStore(unittest.TestCase):
    def setUp(self):
        self.st = S.MemoryStore()
        self.st.bootstrap()

    def test_bootstrap_crea_y_siembra_sin_duplicar(self):
        self.assertEqual(len(self.st.leer(C.HOJA_CATALOGO)), 203)
        self.assertEqual(len(self.st.leer(C.HOJA_PERSONAS)), 5)
        self.assertGreaterEqual(len(self.st.leer(C.HOJA_FERIADOS)), 10)
        self.st.bootstrap()   # segunda vez no duplica
        self.assertEqual(len(self.st.leer(C.HOJA_CATALOGO)), 203)

    def test_bootstrap_agrega_columnas_faltantes(self):
        self.st.hojas[C.HOJA_REGISTROS]["cols"].remove("Nota")
        self.st.bootstrap()
        self.assertIn("Nota", self.st.leer(C.HOJA_REGISTROS).columns)

    def test_alta_edicion_baja_por_id(self):
        filas = [{"ID": S.nuevo_id(), "Fecha": "2026-10-05", "Persona": "Natalia",
                  "Departamento": "IMPUESTOS", "Tarea": "IVA mensual", "Subtarea": "Control",
                  "Minutos": 60, "Nota": "", "Registrado": S.ahora_iso()} for _ in range(3)]
        self.st.agregar(C.HOJA_REGISTROS, filas)
        self.assertEqual(len(self.st.leer(C.HOJA_REGISTROS)), 3)
        self.assertTrue(self.st.actualizar_por_id(C.HOJA_REGISTROS, filas[0]["ID"],
                                                  {"Minutos": 90, "Nota": "ok"}))
        df = self.st.leer(C.HOJA_REGISTROS)
        self.assertEqual(df.loc[df.ID == filas[0]["ID"], "Minutos"].iloc[0], "90")
        self.assertEqual(self.st.borrar_por_id(C.HOJA_REGISTROS, [filas[1]["ID"], filas[2]["ID"]]), 2)
        self.assertEqual(len(self.st.leer(C.HOJA_REGISTROS)), 1)

    def test_flujo_completo_hasta_el_resumen(self):
        filas = calc.generar_registros("Natalia", "IMPUESTOS", "IVA mensual", "Control", "",
                                       date(2026, 10, 5), date(2026, 10, 5), FER, minutos=450)
        for f in filas:
            f.update({"ID": S.nuevo_id(), "Registrado": S.ahora_iso()})
        self.st.agregar(C.HOJA_REGISTROS, filas)
        reg = calc.preparar_registros(self.st.leer(C.HOJA_REGISTROS), self.st.leer(C.HOJA_CATALOGO))
        r = calc.resumen_persona_mes(reg, "Natalia", 2026, 10, HD, FER)
        self.assertEqual((r["trabajo"], r["extra"]), (7.5, 1.5))


if __name__ == "__main__":
    unittest.main()
