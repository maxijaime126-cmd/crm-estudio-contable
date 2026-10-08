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
        self.assertEqual(df["Departamento"].nunique(), 10)      # 9 departamentos + GENERAL
        self.assertEqual(len(df), 179)
        # Disponible y Ausencias existen una sola vez, en GENERAL (no en cada departamento)
        gen = df[df["Departamento"] == C.DEPTO_GENERAL]
        self.assertEqual(set(gen["Tarea"]), {"Disponible", "Ausencias"})
        for dep in set(df["Departamento"]) - {C.DEPTO_GENERAL}:
            tareas = set(df[df["Departamento"] == dep]["Tarea"])
            self.assertNotIn("Disponible", tareas)
            self.assertNotIn("Ausencias", tareas)
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
        self.assertEqual(len(self.st.leer(C.HOJA_CATALOGO)), 179)
        self.assertEqual(len(self.st.leer(C.HOJA_PERSONAS)), 5)
        self.assertGreaterEqual(len(self.st.leer(C.HOJA_FERIADOS)), 10)
        self.st.bootstrap()   # segunda vez no duplica
        self.assertEqual(len(self.st.leer(C.HOJA_CATALOGO)), 179)

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


class TestEtapa2(unittest.TestCase):
    def test_semanas_y_objetivo(self):
        sem = calc.semanas_del_mes(2026, 10)
        self.assertEqual(len(sem), 5)
        self.assertEqual((sem[0][1], sem[0][2]), (date(2026, 10, 1), date(2026, 10, 4)))
        self.assertEqual((sem[-1][1], sem[-1][2]), (date(2026, 10, 26), date(2026, 10, 31)))
        reg = registros([("2026-10-05", "Natalia", "IMPUESTOS", "IVA mensual", "Control", 360),
                         ("2026-10-06", "Natalia", "IMPUESTOS", "Disponible", "", 120)])
        obj, cargadas = calc.objetivo_mes(reg, "Natalia", 2026, 10, HD, FER)
        self.assertEqual((obj, cargadas), (126.0, 8.0))

    def test_meses_atras_cruza_el_anio(self):
        self.assertEqual(calc.meses_atras(2026, 2, 3), [(2025, 12), (2026, 1), (2026, 2)])

    def test_semanal_departamentos(self):
        reg = registros([
            ("2026-10-01", "Natalia", "IMPUESTOS", "IVA mensual", "Control", 360),
            ("2026-10-06", "Natalia", "IMPUESTOS", "IVA mensual", "Control", 600),
            ("2026-10-06", "Athina", "DOCUMENTACIÓN", "Carga de documentación", "Sociedades", 360),
        ])
        hp = {"Natalia": 6.0, "Athina": 6.0}
        h, res = calc.semanal_departamentos(reg, 2026, 10, hp, FER, ["IMPUESTOS", "DOCUMENTACIÓN"])
        s1, s2 = h.columns[0], h.columns[1]
        self.assertEqual(h.loc["IMPUESTOS", s1], 6)
        self.assertEqual(h.loc["IMPUESTOS", s2], 10)
        self.assertEqual(res.loc[s1, "Capacidad del equipo (hs)"], 24)        # jue y vie, 2 personas
        self.assertEqual(res.loc[s2, "Capacidad del equipo (hs)"], 60)        # 5 días, 2 personas
        self.assertEqual(res.loc[s2, "Trabajo total (hs)"], 16)
        self.assertEqual(res.loc[s2, "Libre (hs)"], 44)
        self.assertEqual(res.loc[s2, "Ocupación %"], round(16 / 60 * 100))

    def test_tendencia_y_semanal_persona(self):
        reg = registros([
            ("2026-09-10", "Natalia", "IMPUESTOS", "IVA mensual", "Control", 600),
            ("2026-10-05", "Natalia", "IMPUESTOS", "IVA mensual", "Control", 360),
            ("2026-10-05", "Natalia", "IMPUESTOS", "Disponible", "", 120),   # no es trabajo
        ])
        df, meses = calc.tendencia_departamentos(reg, "Natalia", 2026, 10, 3)
        self.assertEqual(meses, ["Ago 2026", "Sep 2026", "Oct 2026"])
        self.assertEqual(df[df.Mes == "Oct 2026"]["Horas"].sum(), 6)
        self.assertEqual(df[df.Mes == "Sep 2026"]["Horas"].sum(), 10)
        d2, semanas = calc.distribucion_semanal(reg, "Natalia", 2026, 10)
        self.assertEqual(d2["Horas"].sum(), 6)
        self.assertEqual(len(semanas), 5)

    def test_desvio_historico(self):
        reg = registros([
            ("2026-08-10", "Natalia", "IMPUESTOS", "IVA mensual", "Control", 1200),   # 20 hs
            ("2026-09-10", "Natalia", "IMPUESTOS", "IVA mensual", "Control", 600),    # 10 hs
            ("2026-10-05", "Natalia", "IMPUESTOS", "IVA mensual", "Control", 1800),   # 30 hs
            ("2026-10-06", "Natalia", "SUELDOS", "Sindicato", "Control", 120),        # nuevo: 2 hs
        ])
        df, usados = calc.desvio_historico(reg, "Natalia", 2026, 10, "Departamento")
        self.assertEqual(usados, 2)
        imp = df[df.Departamento == "IMPUESTOS"].iloc[0]
        self.assertEqual((imp["Actual (hs)"], imp["Promedio previo (hs)"], imp["Desvío (hs)"]), (30, 15, 15))
        self.assertEqual(imp["Desvío %"], 100)
        sue = df[df.Departamento == "SUELDOS"].iloc[0]
        self.assertEqual(sue["Promedio previo (hs)"], 0)
        self.assertTrue(sue["Desvío %"] != sue["Desvío %"])               # NaN: no hay base para comparar

    def test_desvio_sin_historia(self):
        reg = registros([("2026-10-05", "Natalia", "IMPUESTOS", "IVA mensual", "Control", 600)])
        df, usados = calc.desvio_historico(reg, "Natalia", 2026, 10, "Tarea")
        self.assertEqual(usados, 0)
        self.assertEqual(df.iloc[0]["Actual (hs)"], 10)
        self.assertTrue(df.iloc[0]["Promedio previo (hs)"] != df.iloc[0]["Promedio previo (hs)"])


class TestSinDepartamento(unittest.TestCase):
    def test_disponible_y_ausencia_generales_se_reconocen(self):
        reg = registros([
            ("2026-10-05", "Natalia", "GENERAL", "Disponible", "", 120),
            ("2026-10-06", "Natalia", "GENERAL", "Ausencias", "Vacaciones", 360),
            ("2026-10-07", "Natalia", "IMPUESTOS", "IVA mensual", "Control", 300),
        ])
        self.assertEqual(list(reg["Tipo"]), [C.TIPO_DISPONIBLE, C.TIPO_AUSENCIA, C.TIPO_TRABAJO])
        r = calc.resumen_persona_mes(reg, "Natalia", 2026, 10, HD, FER)
        self.assertEqual((r["disponible"], r["ausencia"], r["trabajo"]), (2, 6, 5))
        self.assertEqual(r["capacidad"], 21 * 6 - 6)

    def test_el_departamento_de_una_hora_libre_no_cambia_los_numeros(self):
        con = registros([("2026-10-05", "Natalia", "CONTABILIDAD", "Disponible", "", 120)])
        sin = registros([("2026-10-05", "Natalia", "GENERAL", "Disponible", "", 120)])
        a = calc.resumen_persona_mes(con, "Natalia", 2026, 10, HD, FER)
        b = calc.resumen_persona_mes(sin, "Natalia", 2026, 10, HD, FER)
        self.assertEqual({k: v for k, v in a.items() if k != "detalle_dias"},
                         {k: v for k, v in b.items() if k != "detalle_dias"})


class TestCalendario(unittest.TestCase):
    def setUp(self):
        self.reg = registros([
            ("2026-10-05", "Natalia", "IMPUESTOS", "IVA mensual", "Control", 840),        # 14 hs
            ("2026-10-05", "Athina", "DOCUMENTACIÓN", "Carga de documentación", "Sociedades", 60),   # 1 h
            ("2026-10-06", "Natalia", "IMPUESTOS", "IVA mensual", "Control", 120),        # 2 hs
            ("2026-10-06", "Athina", "IMPUESTOS", "Disponible", "", 360),                 # solo disponible
        ])
        self.hp = {"Natalia": 6.0, "Athina": 6.0}
        self.deptos = ["IMPUESTOS", "DOCUMENTACIÓN"]
        self.horas, self.cap = calc.matriz_departamentos(self.reg, 2026, 10, self.hp, FER, self.deptos)
        self.datos = calc.dias_con_datos(self.reg, 2026, 10)

    def test_dias_con_datos(self):
        self.assertEqual(self.datos, {date(2026, 10, 5), date(2026, 10, 6)})

    def test_color_contra_capacidad_del_equipo(self):
        z, txt = calc.calendario_z(self.horas, self.cap, self.datos, "equipo")
        d5 = date(2026, 10, 5)
        self.assertEqual(z.index[0], calc.TOTAL_EQUIPO)                       # el total va primero
        self.assertAlmostEqual(z.loc["DOCUMENTACIÓN", d5], 1 / 12 * 100, places=3)   # 1 h de 12 hs: casi nada
        self.assertEqual(z.loc["IMPUESTOS", d5], 100.0)                        # 14 hs sobre 12: tope
        self.assertEqual(z.loc[calc.TOTAL_EQUIPO, d5], 100.0)
        self.assertEqual(txt.loc["IMPUESTOS", d5], "14.0")

    def test_dias_sin_carga_quedan_sin_datos(self):
        z, _ = calc.calendario_z(self.horas, self.cap, self.datos, "equipo")
        sin = date(2026, 10, 7)
        self.assertTrue(z[sin].isna().all())                                   # nadie cargó: gris, no verde
        self.assertEqual(z.loc["DOCUMENTACIÓN", date(2026, 10, 6)], 0.0)       # día con datos y 0 hs: verde

    def test_color_contra_el_pico_del_departamento(self):
        z, _ = calc.calendario_z(self.horas, self.cap, self.datos, "pico")
        self.assertEqual(z.loc["DOCUMENTACIÓN", date(2026, 10, 5)], 100.0)     # su único día fuerte
        self.assertAlmostEqual(z.loc["IMPUESTOS", date(2026, 10, 6)], 2 / 14 * 100, places=3)

    def test_fin_de_semana_trabajado_es_todo_extra(self):
        reg = registros([("2026-10-05", "Natalia", "IMPUESTOS", "IVA mensual", "Control", 60),
                         ("2026-10-10", "Natalia", "IMPUESTOS", "IVA mensual", "Control", 120)])
        h, c = calc.matriz_departamentos(reg, 2026, 10, {"Natalia": 6.0}, FER, ["IMPUESTOS"])
        z, _ = calc.calendario_z(h, c, calc.dias_con_datos(reg, 2026, 10), "equipo")
        self.assertEqual(z.loc["IMPUESTOS", date(2026, 10, 10)], 100.0)

    def test_grilla_de_un_departamento(self):
        etiquetas, z, txt = calc.grilla_calendario(self.horas.loc["IMPUESTOS"], self.cap, self.datos,
                                                   FER, 2026, 10, "equipo")
        self.assertEqual(len(etiquetas), 5)
        self.assertTrue(all(len(f) == 7 for f in z))
        self.assertEqual(txt[0][0], "")                                        # lunes 28/9: fuera del mes
        self.assertEqual(txt[0][3], "<b>1</b>")              # jueves 1/10, sin horas
        self.assertEqual(txt[1][0], "<b>5</b><br>14.0 hs")                     # lunes 5/10
        self.assertEqual(z[1][0], 100.0)
        self.assertEqual(txt[2][0], "12<br>feriado")                           # lunes 12/10
        self.assertTrue(z[2][0] != z[2][0])                                    # el feriado es gris

    def test_quien_puede_ayudar_distingue_sin_carga(self):
        r = calc.quien_puede_ayudar(self.reg, date(2026, 10, 7), self.hp, FER)
        self.assertTrue((r["Estado"] == "sin carga ese día").all())
        self.assertTrue(r["Margen (hs)"].isna().all())                         # no se asume que están libres
        r = calc.quien_puede_ayudar(self.reg, date(2026, 10, 10), self.hp, FER)
        self.assertTrue((r["Estado"] == "día no hábil").all())
        r = calc.quien_puede_ayudar(self.reg, date(2026, 10, 6), self.hp, FER)
        self.assertEqual(r.iloc[0]["Persona"], "Athina")
        self.assertEqual(r.iloc[0]["Margen (hs)"], 6.0)


class TestResumenDiario(unittest.TestCase):
    def test_estados_por_dia(self):
        reg = registros([
            ("2026-10-05", "Natalia", "IMPUESTOS", "IVA mensual", "Control", 300),      # 5 hs trabajo
            ("2026-10-05", "Natalia", "GENERAL", "Disponible", "", 60),                 # + 1 h libre = completo
            ("2026-10-06", "Natalia", "IMPUESTOS", "IVA mensual", "Control", 180),      # 3 hs: faltan 3
            ("2026-10-07", "Natalia", "IMPUESTOS", "IVA mensual", "Control", 420),      # 7 hs: 1 de más
            ("2026-10-10", "Natalia", "IMPUESTOS", "IVA mensual", "Control", 60),       # sábado
            ("2026-10-13", "Natalia", "GENERAL", "Ausencias", "Vacaciones", 360),       # planificada a futuro
        ])
        r = calc.resumen_diario(reg, "Natalia", 2026, 10, HD, FER, date(2026, 10, 9))
        est = dict(zip(r["Fecha"], r["Estado"]))
        self.assertEqual(est[date(2026, 10, 5)], "✅ completo")
        self.assertEqual(est[date(2026, 10, 6)], "⚠️ faltan 3,0 hs")
        self.assertEqual(est[date(2026, 10, 7)], "⏱️ 1,0 hs de más")
        self.assertEqual(est[date(2026, 10, 8)], "❌ sin carga")
        self.assertNotIn(date(2026, 10, 9), est)                       # hoy todavía se puede cargar
        self.assertEqual(est[date(2026, 10, 10)], "⏱️ día no hábil: cuenta como extra")
        self.assertNotIn(date(2026, 10, 12), est)                      # feriado futuro sin carga
        self.assertEqual(est[date(2026, 10, 13)], "✅ completo")       # la vacación ya cargada cuenta
        fila5 = r[r["Fecha"] == date(2026, 10, 5)].iloc[0]
        self.assertEqual((fila5["Trabajo (hs)"], fila5["Disponible (hs)"], fila5["Total (hs)"]), (5.0, 1.0, 6.0))


class TestEstadoDelDia(unittest.TestCase):
    def test_quien_completo_el_dia(self):
        reg = registros([
            ("2026-10-05", "Natalia", "IMPUESTOS", "IVA mensual", "Control", 300),
            ("2026-10-05", "Natalia", "GENERAL", "Disponible", "", 60),                 # completo
            ("2026-10-05", "Athina", "IMPUESTOS", "IVA mensual", "Control", 120),       # faltan 4
            ("2026-10-05", "Johana", "GENERAL", "Ausencias", "Vacaciones", 360),        # ausente todo el día
        ])
        hp = {"Natalia": 6.0, "Athina": 6.0, "Maximiliano": 6.0, "Johana": 6.0}
        t = calc.estado_dia(reg, date(2026, 10, 5), hp, FER)
        est = dict(zip(t["Persona"], t["Estado"]))
        self.assertEqual(est["Natalia"], "✅ completo")
        self.assertEqual(est["Athina"], "⚠️ faltan 4,0 hs")
        self.assertEqual(est["Maximiliano"], "❌ sin carga")
        self.assertEqual(est["Johana"], "✅ completo")                    # la vacación cuenta como día cubierto
        self.assertEqual(int(t["Estado"].str.startswith("✅").sum()), 2)

    def test_dia_no_habil(self):
        reg = registros([("2026-10-12", "Natalia", "IMPUESTOS", "IVA mensual", "Control", 60)])
        t = calc.estado_dia(reg, date(2026, 10, 12), {"Natalia": 6.0, "Athina": 6.0}, FER)
        est = dict(zip(t["Persona"], t["Estado"]))
        self.assertEqual(est["Natalia"], "⏱️ día no hábil: cuenta como extra")
        self.assertEqual(est["Athina"], "— día no hábil")


class FakeWS:
    def __init__(self, title, ids):
        self.title, self.rows, self.col_count, self.id = title, [], 6, next(ids)

    def row_values(self, n):
        return list(self.rows[0]) if self.rows else []

    def col_values(self, n):
        return [r[n - 1] if len(r) >= n else "" for r in self.rows]

    def update(self, range_name, values, **k):
        assert range_name == "A1"
        if self.rows:
            self.rows[0] = list(values[0])
        else:
            self.rows.append(list(values[0]))

    def resize(self, cols=None, **k):
        self.col_count = cols

    def append_rows(self, values, **k):
        self.rows.extend([list(v) for v in values])

    def get_all_values(self):
        return self.rows


class FakeSS:
    def __init__(self, carrera=()):
        import itertools
        self.ids = itertools.count(1)
        self.hojas, self.carrera = [], set(carrera)

    def worksheets(self):
        return list(self.hojas)

    def worksheet(self, t):
        return next(w for w in self.hojas if w.title == t)

    def add_worksheet(self, title, rows, cols):
        existe = any(w.title.lower() == title.lower() for w in self.hojas)
        if title in self.carrera:           # otra sesión la crea justo antes
            self.carrera.discard(title)
            self.hojas.append(FakeWS(title, self.ids))
            existe = True
        if existe:
            raise Exception('APIError: [400]: Invalid requests[0].addSheet: A sheet with the name '
                            f'"{title}" already exists. Please enter another name.')
        w = FakeWS(title, self.ids)
        self.hojas.append(w)
        return w


class TestBootstrapSheets(unittest.TestCase):
    def test_hoja_existente_con_otras_mayusculas_y_datos_propios(self):
        ss = FakeSS()
        fer = FakeWS("feriados", ss.ids)
        fer.rows = [["fecha", "motivo"], ["01/01/2027", "Año Nuevo"]]
        ss.hojas.append(fer)
        S.SheetsStore(ss).bootstrap()
        self.assertEqual(len([w for w in ss.hojas if w.title.lower() == "feriados"]), 1)
        self.assertEqual(fer.rows[0], ["Fecha", "Motivo"])      # normaliza el encabezado
        self.assertEqual(len(fer.rows), 2)                       # no pisa ni siembra: ya tenía datos
        self.assertEqual({w.title for w in ss.hojas},
                         {"Catalogo", "Personas", "Registros", "feriados"})

    def test_carrera_entre_sesiones(self):
        ss = FakeSS(carrera={"Feriados"})
        S.SheetsStore(ss).bootstrap()
        self.assertEqual(len([w for w in ss.hojas if w.title.lower() == "feriados"]), 1)
        fer = next(w for w in ss.hojas if w.title == "Feriados")
        self.assertEqual(fer.rows[0], ["Fecha", "Motivo"])
        self.assertGreater(len(fer.rows), 5)                     # se sembraron los feriados

    def test_arranque_repetido_no_duplica(self):
        ss = FakeSS()
        S.SheetsStore(ss).bootstrap()
        n = {w.title: len(w.rows) for w in ss.hojas}
        S.SheetsStore(ss).bootstrap()
        self.assertEqual(n, {w.title: len(w.rows) for w in ss.hojas})
        self.assertEqual(n["Catalogo"], 180)                     # encabezado + 179 filas

    def test_agrega_columnas_faltantes_sin_mover_las_existentes(self):
        ss = FakeSS()
        reg = FakeWS("Registros", ss.ids)
        reg.rows = [["Fecha", "ID", "Minutos"], ["2026-10-05", "a1", "60"]]
        ss.hojas.append(reg)
        S.SheetsStore(ss).bootstrap()
        self.assertEqual(reg.rows[0][:3], ["Fecha", "ID", "Minutos"])
        self.assertIn("Persona", reg.rows[0])


if __name__ == "__main__":
    unittest.main()
