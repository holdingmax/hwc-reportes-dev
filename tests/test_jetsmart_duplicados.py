"""Filas repetidas IDENTICAS en la rendicion de Ariel (src/jetsmart_cruce.py).

Correr con:
    python -m unittest tests.test_jetsmart_duplicados -v

Regla de negocio: una fila repetida con TODAS las columnas iguales es la
misma guia cargada dos veces: se cuenta una vez y se avisa. Una guia repetida
con algun dato distinto NO se toca (revision manual): ejemplo real, la 97479
de agosto 2026 (el destinatario dice "LIBERBOX" en una fila y "LIBERBOX SAS"
en la otra; aunque el dinero sea igual, la decision es de una persona).

Los tests con archivos reales (octubre 2026: 20 guias del 31/10 repetidas)
usan tests/fixtures/jetsmart/ (no versionado) y se saltean si faltan.
"""

import io
import unittest
from pathlib import Path

import openpyxl
import pandas as pd

from src.jetsmart_archivos import TIPO_ARIEL, cargar_guias_sistema
from src.jetsmart_builder import build_jetsmart_guias, build_jetsmart_resumen, write_jetsmart_liquidacion
from src.jetsmart_cruce import (
    _AR_COLUMNS,
    EXC_DUPLICADA,
    _limpiar_ariel_detalle,
    cruzar_jetsmart,
    leer_ariel,
    load_ariel_raw,
)

FIX = Path(__file__).resolve().parent / "fixtures" / "jetsmart"
GUIAS_OCT = FIX / "Guias_Aerea_10-2026.xlsx"
ARIEL_OCT = FIX / "Ariel_Rendicion_octubre_2026.xlsx"
ARIEL_AGO = FIX / "Ariel_Rendicion_agosto_2026.xlsx"

FILA = [827, 50200474, 94542, pd.Timestamp(2026, 10, 31), 3039, "AMDM S.R.L", "JULIO BURGOA", "AEP", "BRC", 1, 8, 8, 736, 10000, "ART"]


def _tabla(*filas) -> pd.DataFrame:
    return pd.DataFrame(list(filas), columns=_AR_COLUMNS)


class RegladeRepetidasIdenticas(unittest.TestCase):
    def test_fila_identica_se_cuenta_una_vez_y_se_anota(self):
        leido = _limpiar_ariel_detalle(_tabla(FILA, FILA, FILA[:2] + [94543] + FILA[3:]), "Venta")
        self.assertEqual(sorted(leido.guias["Nro de guía"]), [94542, 94543])
        self.assertEqual(leido.repetidas_identicas, [94542])

    def test_misma_guia_con_un_dato_distinto_no_se_toca(self):
        # Igual que 97479 de agosto: solo cambia el destinatario.
        distinta = list(FILA)
        distinta[6] = "JULIO BURGOA SA"
        leido = _limpiar_ariel_detalle(_tabla(FILA, distinta), "Venta")
        self.assertEqual(len(leido.guias), 2)
        self.assertEqual(leido.repetidas_identicas, [])
        # ... y el cruce la manda a revision manual como siempre.
        r = cruzar_jetsmart(_export_con(94542), leido.guias, ("OCT", "26"), None, None)
        self.assertEqual(set(r.excepciones.loc[r.excepciones["Tipo"] == EXC_DUPLICADA, "Nro guía"]), {94542})
        self.assertEqual(len(r.incluidas), 0)

    def test_anuladas_no_cuentan_como_repetidas(self):
        # Ariel reusa el numero de una anulada para otra guia valida: no es un duplicado.
        anulada = [827, 50201502, 94542, "ANULADA"] + [None] * 11
        leido = _limpiar_ariel_detalle(_tabla(FILA, anulada), "Venta")
        self.assertEqual((len(leido.guias), leido.anuladas, leido.repetidas_identicas), (1, 1, []))

    def test_sin_repetidas_el_resumen_lo_dice(self):
        leido = _limpiar_ariel_detalle(_tabla(FILA), "Venta")
        r = cruzar_jetsmart(_export_con(94542), leido.guias, ("OCT", "26"), None, None, ariel_repetidas_identicas=leido.repetidas_identicas)
        self.assertEqual(r.resumen["ariel_repetidas_identicas"], [])


def _export_con(guia: int) -> pd.DataFrame:
    from src.config import (COL_JS_CLIENTE, COL_JS_CREACION, COL_JS_DESTINO, COL_JS_EMPRESA, COL_JS_ESTADO,
                            COL_JS_GUIA, COL_JS_KGS, COL_JS_ORIGEN, COL_JS_PRIORIDAD)
    return pd.DataFrame([{
        COL_JS_CREACION: "31/10/2026", COL_JS_GUIA: guia, COL_JS_CLIENTE: "AMDM S.R.L", COL_JS_ORIGEN: "AEP",
        COL_JS_DESTINO: "BRC", COL_JS_ESTADO: "Entregada", COL_JS_KGS: 8.0, COL_JS_PRIORIDAD: 10000.0, COL_JS_EMPRESA: "Jetsmart",
    }])


@unittest.skipUnless(GUIAS_OCT.exists() and ARIEL_OCT.exists(), "faltan los archivos de octubre")
class OctubreReal(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.leido = leer_ariel(ARIEL_OCT, cruda=True)
        cls.export = cargar_guias_sistema(GUIAS_OCT)
        cls.cruce = cruzar_jetsmart(
            cls.export, cls.leido.guias, ("OCT", "26"), None, None, ariel_repetidas_identicas=cls.leido.repetidas_identicas,
        )

    def test_20_guias_repetidas_del_31_10(self):
        self.assertEqual(len(self.leido.repetidas_identicas), 20)
        self.assertEqual(len(self.leido.guias), 1953)
        self.assertEqual(self.leido.anuladas, 8)
        self.assertFalse(self.leido.guias["Nro de guía"].duplicated().any())
        fechas = self.leido.guias[self.leido.guias["Nro de guía"].isin(self.leido.repetidas_identicas)]["Fecha de guia"]
        self.assertEqual(set(fechas.dt.strftime("%d/%m/%Y")), {"31/10/2026"})

    def test_load_ariel_raw_devuelve_lo_mismo_sin_repetidas(self):
        guias, anuladas = load_ariel_raw(ARIEL_OCT)
        self.assertEqual((len(guias), anuladas), (1953, 8))

    def test_las_20_ya_no_van_a_revision_y_entran_a_la_liquidacion(self):
        self.assertNotIn(EXC_DUPLICADA, set(self.cruce.excepciones["Tipo"]))
        repetidas = set(self.leido.repetidas_identicas)
        self.assertEqual(len(repetidas), 20)
        # Las que estan en las guias del sistema (matchean) entran con los valores de Ariel.
        en_sistema = repetidas & set(pd.to_numeric(self.export["# Guía"]))
        self.assertTrue(en_sistema <= set(self.cruce.incluidas["# Guía"]))
        self.assertEqual(self.cruce.resumen["ariel_repetidas_identicas"], sorted(self.leido.repetidas_identicas))

    def test_cruce_con_guias_del_sistema_sin_diferencias(self):
        # 1.934 en comun, 0 diferencias de kg ni de ingreso (medido a mano antes de escribir el codigo).
        self.assertNotIn("Diferencia de Kg / ingreso", set(self.cruce.excepciones["Tipo"]))
        comunes = set(pd.to_numeric(self.export["# Guía"])) & set(self.leido.guias["Nro de guía"])
        self.assertEqual(len(comunes), 1934)

    def test_el_excel_lleva_el_aviso(self):
        guias = build_jetsmart_guias(self.cruce.incluidas)
        resumen = build_jetsmart_resumen(guias, 1485.0, {"EZE": 0, "MDZ": 0})
        buf = io.BytesIO()
        write_jetsmart_liquidacion(
            guias, resumen, ("OCT", "26"), buf,
            cruce={"resumen": self.cruce.resumen, "excepciones": self.cruce.excepciones,
                   "incluidas_ultimo_dia": self.cruce.incluidas_ultimo_dia},
        )
        buf.seek(0)
        wb = openpyxl.load_workbook(buf)
        self.assertIn("REPETIDAS EN ARIEL", wb.sheetnames)
        self.assertEqual(wb["REPETIDAS EN ARIEL"].max_row, 21)  # encabezado + 20 guias


@unittest.skipUnless(ARIEL_AGO.exists(), "falta la rendicion de agosto")
class AgostoNoCambia(unittest.TestCase):
    def test_97479_sigue_siendo_revision_manual(self):
        # En agosto la guia repetida NO es identica (cambia el destinatario): no se descarta.
        leido = leer_ariel(ARIEL_AGO, cruda=True)
        self.assertEqual(leido.repetidas_identicas, [])
        self.assertEqual(int((leido.guias["Nro de guía"] == 97479).sum()), 2)


if __name__ == "__main__":
    unittest.main()
