"""Validacion del cruce JetSmart (export de guias vs Ariel) contra las
liquidaciones finales reales de Anita: hoja GUIAS de LIQ_ECS de junio (1925
guias) y julio (1949 guias) 2026.

Correr con:
    python -m unittest tests.test_jetsmart_cruce -v

Necesita en tests/fixtures/jetsmart/ (no se versionan, son datos del cliente):
LIQUIDACION_06-2026.xlsx, LIQUIDACION_07-2026.xlsx (archivos de trabajo, con
hojas BD y ARIEL) y LIQ_ECS_06/07-2026_00000005.xlsx (liquidaciones finales).

Criterio: ninguna decision automatica del cruce contradice a Anita (lo que
el cruce incluye esta en la final; lo que excluye sin marcar, no), y todo lo
que difiere queda marcado como excepcion. Diferencias conocidas y aceptadas
(ajustes manuales de Anita que no estan en ningun archivo):
- julio: 94940 tiene 7.956 en Ariel y 12.000 en la final (corregida a mano).
- junio: 92364, 92378 y 92409 (del 30/05) agregadas a mano a la final.
- junio: no hay archivos de mayo, asi que las guias del 01-02/06 que Ariel
  declara y que estarian en el export de mayo quedan a revision manual.
"""

import unittest
from pathlib import Path

import openpyxl
import pandas as pd

from src.jetsmart_builder import (
    build_jetsmart_guias,
    build_jetsmart_resumen,
    detect_jetsmart_period,
    load_jetsmart_export,
)
from src.jetsmart_cruce import (
    EXC_DIFERENCIA,
    EXC_PENDIENTE,
    EXC_REVISION,
    INC_ULTIMO_DIA,
    cruzar_jetsmart,
    load_ariel,
)

FIX = Path(__file__).resolve().parent / "fixtures" / "jetsmart"
TRABAJO_JUN, FINAL_JUN = FIX / "LIQUIDACION_06-2026.xlsx", FIX / "LIQ_ECS_06-2026_00000005.xlsx"
TRABAJO_JUL, FINAL_JUL = FIX / "LIQUIDACION_07-2026.xlsx", FIX / "LIQ_ECS_07-2026_00000005.xlsx"
HAY_FIXTURES = all(p.exists() for p in (TRABAJO_JUN, FINAL_JUN, TRABAJO_JUL, FINAL_JUL))


def _final(path) -> set[int]:
    g = pd.read_excel(path, sheet_name="GUIAS ", usecols=range(15))
    return set(g.loc[g["AWB"].notna(), "Nro guía"].astype(int))


def _excepciones(resultado, tipo) -> set[int]:
    exc = resultado.excepciones
    return set(exc.loc[exc["Tipo"] == tipo, "Nro guía"])


@unittest.skipUnless(HAY_FIXTURES, "faltan los archivos de tests/fixtures/jetsmart/")
class TestCruce(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.export_jun = load_jetsmart_export(TRABAJO_JUN)
        cls.export_jul = load_jetsmart_export(TRABAJO_JUL)
        cls.ariel_jun, cls.anuladas_jun = load_ariel(TRABAJO_JUN)
        cls.ariel_jul, cls.anuladas_jul = load_ariel(TRABAJO_JUL)
        cls.final_jun, cls.final_jul = _final(FINAL_JUN), _final(FINAL_JUL)
        cls.periodo_jun, _ = detect_jetsmart_period(cls.export_jun)
        cls.periodo_jul, _ = detect_jetsmart_period(cls.export_jul)

        # Julio: como al cierre real -- sin el export de agosto.
        cls.jul = cruzar_jetsmart(cls.export_jul, cls.ariel_jul, cls.periodo_jul, cls.export_jun, cls.ariel_jun)
        # Junio: sin archivos de mayo; con y sin el export de julio.
        cls.jun_con_sig = cruzar_jetsmart(cls.export_jun, cls.ariel_jun, cls.periodo_jun, None, None, cls.export_jul)
        cls.jun_sin_sig = cruzar_jetsmart(cls.export_jun, cls.ariel_jun, cls.periodo_jun, None, None)

    def test_anuladas_se_descartan_antes_de_cruzar(self):
        self.assertEqual((self.anuladas_jun, self.anuladas_jul), (30, 23))
        # Ariel de julio reusa 15 numeros de guias anuladas con otro AWB:
        # sin las anuladas no queda ningun duplicado.
        self.assertFalse(self.ariel_jul["Nro de guía"].duplicated().any())

    # --- julio ---
    def test_julio_incluidas_igual_a_la_final(self):
        self.assertEqual(set(self.jul.incluidas["# Guía"]), self.final_jul)
        self.assertEqual(len(self.jul.incluidas), 1949)

    def test_julio_excepciones(self):
        self.assertEqual(len(self.jul.excepciones), 2)
        self.assertEqual(_excepciones(self.jul, EXC_DIFERENCIA), {94940})
        self.assertEqual(_excepciones(self.jul, EXC_PENDIENTE), {96640})

    def test_julio_ya_declaradas_en_junio_no_son_pendientes(self):
        # 35 guias del 29-30/06 que estan en el export de julio pero que
        # Ariel declaro (y Anita liquido) en junio.
        self.assertEqual(self.jul.resumen["ya_declaradas_mes_anterior"], 35)
        self.assertFalse(set(self.jul.incluidas["# Guía"]) & (self.final_jun - {94582, 94585, 94587}))

    def test_julio_regla_ultimo_dia_y_pendientes_de_junio(self):
        motivos = self.jul.resumen["motivos_inclusion"]
        self.assertEqual(motivos[INC_ULTIMO_DIA], 21)
        self.assertEqual(len(self.jul.incluidas_ultimo_dia), 21)
        self.assertTrue((self.jul.incluidas_ultimo_dia["Fecha Ariel"] == "31/07/2026").all())
        # Pendientes de declarar de junio que Ariel declara en julio.
        self.assertEqual(motivos["pendiente del mes anterior, declarada ahora"], 3)

    def test_julio_totales_contra_la_final(self):
        guias = build_jetsmart_guias(self.jul.incluidas)
        resumen = build_jetsmart_resumen(guias, 1485, {"EZE": 56, "MDZ": 38})
        ref = openpyxl.load_workbook(FINAL_JUL, data_only=True).worksheets[1]
        # Unica diferencia: la 94940 (Ariel 7.956 vs 12.000 corregida a mano).
        self.assertAlmostEqual(resumen["ventas_netas"] - ref["B6"].value, 7956 - 12000, places=4)
        self.assertAlmostEqual(resumen["gha_services"], ref["B9"].value, places=4)
        self.assertAlmostEqual(resumen["comision_inter"], ref["B8"].value, places=4)

    # --- junio ---
    def test_junio_nada_incluido_de_mas(self):
        for resultado in (self.jun_con_sig, self.jun_sin_sig):
            with self.subTest(export_siguiente=resultado.resumen["export_siguiente_cargado"]):
                self.assertEqual(set(resultado.incluidas["# Guía"]) - self.final_jun, set())

    def test_junio_lo_que_falta_esta_marcado_o_es_ajuste_manual(self):
        sin_mayo = {92373, 92375, 92380, 92381, 92406, 92535}
        manuales_anita = {92364, 92378, 92409}
        for resultado in (self.jun_con_sig, self.jun_sin_sig):
            with self.subTest(export_siguiente=resultado.resumen["export_siguiente_cargado"]):
                faltan = self.final_jun - set(resultado.incluidas["# Guía"])
                self.assertEqual(faltan, sin_mayo | manuales_anita)
                self.assertLessEqual(sin_mayo, _excepciones(resultado, EXC_REVISION))

    def test_junio_excepciones(self):
        r = self.jun_con_sig
        self.assertEqual(len(r.excepciones), 17)
        self.assertEqual(_excepciones(r, EXC_DIFERENCIA), {92542, 93172, 93477})
        self.assertEqual(_excepciones(r, EXC_PENDIENTE), {93960, 94582, 94585, 94587})
        # Las 3 del 10/06 que no aparecen en ningun lado (Anita tampoco las
        # liquido) quedan a revision manual, no incluidas.
        self.assertLessEqual({93051, 93097, 93115, 91811}, _excepciones(r, EXC_REVISION))

    def test_junio_ultimo_dia_confirmado_con_export_siguiente_o_por_regla(self):
        # Con el export de julio, las 35 del 30/06 quedan confirmadas; sin
        # el, entran igual por la regla de ultimo dia (con aviso).
        self.assertEqual(self.jun_con_sig.resumen["motivos_inclusion"]["confirmada con el export del mes siguiente"], 35)
        self.assertTrue(self.jun_con_sig.incluidas_ultimo_dia.empty)
        self.assertEqual(len(self.jun_sin_sig.incluidas_ultimo_dia), 35)
        self.assertEqual(set(self.jun_con_sig.incluidas["# Guía"]), set(self.jun_sin_sig.incluidas["# Guía"]))

    def test_diferencias_usan_los_valores_de_ariel(self):
        inc = self.jun_con_sig.incluidas.set_index("# Guía")
        self.assertEqual(inc.loc[93172, "KGs"], 956)          # export dice 1035
        self.assertEqual(inc.loc[92542, "$ Prioridad"], 166717.4)  # export dice 166717.44


if __name__ == "__main__":
    unittest.main()
