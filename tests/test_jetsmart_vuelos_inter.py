"""Comisiones por vuelos internacionales de JetSmart: el importe por vuelo es
un dato de la liquidacion (no una constante), se guarda con ella y el
Historial y el Excel usan el que se uso, no el de config.py.

Correr con:
    python -m unittest tests.test_jetsmart_vuelos_inter -v

Guias sinteticas: no necesita los archivos reales ni la base.
"""

import io
import unittest
from unittest import mock

import pandas as pd
from openpyxl import load_workbook

from src import jetsmart_builder
from src.jetsmart_builder import (
    COL_GUIAS_IVA,
    COL_GUIAS_KG,
    COL_GUIAS_ORIGEN,
    COL_GUIAS_TOTAL,
    build_jetsmart_resumen,
    build_jetsmart_resumen_guardado,
    write_jetsmart_liquidacion,
)

TC = 1000.0


def _guias() -> pd.DataFrame:
    return pd.DataFrame({COL_GUIAS_ORIGEN: ["AEP"], COL_GUIAS_KG: [10.0], COL_GUIAS_TOTAL: [1210.0], COL_GUIAS_IVA: [210.0]})


def _por_estacion(resumen: dict) -> dict:
    return {c["estacion"]: c for c in resumen["comisiones_inter"]}


class ImportePorVuelo(unittest.TestCase):
    def test_sin_importes_rige_config(self):
        c = _por_estacion(build_jetsmart_resumen(_guias(), TC, {"EZE": 2, "MDZ": 1}))
        self.assertEqual((c["EZE"]["usd_por_vuelo"], c["MDZ"]["usd_por_vuelo"]), (17.5, 35.0))
        self.assertEqual(c["EZE"]["pesos"], 2 * 17.5 * TC)

    def test_importes_de_la_liquidacion_pisan_config(self):
        resumen = build_jetsmart_resumen(_guias(), TC, {"EZE": 49, "MDZ": 10}, {"EZE": 20.25, "MDZ": 40.0})
        c = _por_estacion(resumen)
        self.assertEqual(c["EZE"]["usd_por_vuelo"], 20.25)
        self.assertEqual(c["EZE"]["pesos"], 49 * 20.25 * TC)
        self.assertEqual(c["MDZ"]["pesos"], 10 * 40.0 * TC)
        self.assertEqual(resumen["comision_inter"], -(49 * 20.25 + 10 * 40.0) * TC)

    def test_importe_negativo_se_rechaza(self):
        with self.assertRaises(ValueError):
            build_jetsmart_resumen(_guias(), TC, {"EZE": 1}, {"EZE": -1.0})


class HistorialUsaLoGuardado(unittest.TestCase):
    PARAMS = {"tipo_cambio": TC, "vuelos_inter": {"EZE": 49, "MDZ": 10}, "usd_por_vuelo": {"EZE": 20.0, "MDZ": 40.0}}

    def test_recalcula_con_los_importes_guardados_aunque_config_cambie(self):
        original = build_jetsmart_resumen(_guias(), TC, self.PARAMS["vuelos_inter"], self.PARAMS["usd_por_vuelo"])
        with mock.patch.object(jetsmart_builder, "JETSMART_COMISION_INTER_USD_POR_VUELO", {"EZE": 99.0, "MDZ": 99.0}):
            recalculado = build_jetsmart_resumen_guardado(_guias(), self.PARAMS)
        self.assertEqual(recalculado["comision_inter"], original["comision_inter"])
        self.assertEqual(recalculado["total_wcs"], original["total_wcs"])

    def test_liquidacion_vieja_sin_importes_usa_config(self):
        vieja = {"tipo_cambio": TC, "vuelos_inter": {"EZE": 56, "MDZ": 38}}
        c = _por_estacion(build_jetsmart_resumen_guardado(_guias(), vieja))
        self.assertEqual((c["EZE"]["usd_por_vuelo"], c["MDZ"]["usd_por_vuelo"]), (17.5, 35.0))

    def test_excel_trae_los_importes_usados(self):
        resumen = build_jetsmart_resumen_guardado(_guias(), self.PARAMS)
        buffer = io.BytesIO()
        write_jetsmart_liquidacion(_guias(), resumen, ("OCT", "26"), buffer)
        ws = load_workbook(io.BytesIO(buffer.getvalue()))["COMISIONES INTER"]
        filas = {r[0]: r for r in ws.iter_rows(min_row=2, max_row=3, values_only=True)}
        self.assertEqual(filas["EZE"][1:3], (49, 20.0))
        self.assertEqual(filas["MDZ"][1:3], (10, 40.0))


if __name__ == "__main__":
    unittest.main()
