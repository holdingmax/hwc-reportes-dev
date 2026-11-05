"""Historial de JetSmart: el numero de guia se ve entero, sin decimales.

Correr con:
    python -m unittest tests.test_jetsmart_historial_guias -v

Sin archivos reales ni base: simula lo que devuelve la base y lo muestra con
la pantalla real (AppTest).
"""

import io
import os
import unittest
from datetime import datetime
from pathlib import Path
from unittest import mock

import pandas as pd
from openpyxl import load_workbook
from streamlit.testing.v1 import AppTest

from src.jetsmart_builder import (
    COL_GUIAS_GUIA,
    build_jetsmart_resumen_guardado,
    guias_guardadas,
    write_jetsmart_liquidacion,
)

APP = Path(__file__).resolve().parent.parent / "app.py"

# Como vuelven de la base: algunos numeros con ruido decimal, otros enteros.
REGISTROS = [
    {"Nro guía": 102176.99999999999, "Kg": 10.0, "GRAVADO": 1000.0, "IVA": 210.0, "NO GRAV": 0.0, "TOTAL": 1210.0, "Origen": "AEP"},
    {"Nro guía": 102177, "Kg": 5.5, "GRAVADO": 500.0, "IVA": 105.0, "NO GRAV": 0.0, "TOTAL": 605.0, "Origen": "COR"},
    {"Nro guía": 1234567.0, "Kg": 1.0, "GRAVADO": 100.0, "IVA": 21.0, "NO GRAV": 0.0, "TOTAL": 121.0, "Origen": "AEP"},
]


class GuiasGuardadas(unittest.TestCase):
    def test_numero_de_guia_entero(self):
        guias = guias_guardadas(REGISTROS)
        self.assertEqual(list(guias[COL_GUIAS_GUIA]), [102177, 102177, 1234567])
        self.assertEqual(str(guias[COL_GUIAS_GUIA].dtype), "Int64")

    def test_no_cambia_ningun_otro_dato(self):
        guias = guias_guardadas(REGISTROS)
        esperado = pd.DataFrame(REGISTROS).drop(columns=[COL_GUIAS_GUIA])
        pd.testing.assert_frame_equal(guias.drop(columns=[COL_GUIAS_GUIA]), esperado)

    def test_el_excel_del_historial_trae_la_guia_entera(self):
        guias = guias_guardadas(REGISTROS)
        resumen = build_jetsmart_resumen_guardado(guias, {"tipo_cambio": 1000.0, "vuelos_inter": {}})
        buffer = io.BytesIO()
        write_jetsmart_liquidacion(guias, resumen, ("OCT", "26"), buffer)
        ws = load_workbook(io.BytesIO(buffer.getvalue()))["GUIAS"]
        valores = [ws.cell(row=r, column=1).value for r in (2, 3, 4)]
        self.assertEqual(valores, [102177, 102177, 1234567])
        self.assertTrue(all(isinstance(v, int) for v in valores))


class HistorialEnPantalla(unittest.TestCase):
    """La pantalla real del Historial, con la base simulada."""

    def test_la_tabla_de_guias_muestra_enteros(self):
        fila = pd.DataFrame([{
            "id": 7, "carga_id": 7, "aerolinea": "jetsmart", "estacion": "AR", "tipo_cargo": "jetsmart_liquidacion",
            "periodo_mes": "OCT", "periodo_anio": "26", "periodo": datetime(2026, 10, 1), "cantidad_filas": 3,
            "monto_total": 1000.0, "es_vigente": True, "generado_en": pd.Timestamp("2026-11-04 12:00"),
        }])
        detalle = {"guias": REGISTROS, "parametros": {"tipo_cambio": 1000.0, "vuelos_inter": {}}, "cruce": None}
        env = mock.patch.dict(os.environ)
        env.start()
        os.environ.pop("DATABASE_URL", None)
        self.addCleanup(env.stop)
        for nombre, valor in (
            ("obtener_filtros_historial", {"aerolineas": ["jetsmart"], "periodos": [(datetime(2026, 10, 1), "OCT", "26")], "estaciones": ["AR"]}),
            ("obtener_historial", fila),
            ("obtener_detalle_totales", detalle),
        ):
            patch = mock.patch(f"src.db.{nombre}", return_value=valor)
            patch.start()
            self.addCleanup(patch.stop)
        at = AppTest.from_file(str(APP), default_timeout=120)
        at.run()
        self.assertEqual(len(at.exception), 0, [e.value for e in at.exception])
        tablas = [df.value for df in at.dataframe if COL_GUIAS_GUIA in df.value.columns]
        self.assertTrue(tablas, "no aparecio la tabla de guias del Historial")
        self.assertEqual(list(tablas[0][COL_GUIAS_GUIA]), ["102177", "102177", "1234567"])


if __name__ == "__main__":
    unittest.main()
