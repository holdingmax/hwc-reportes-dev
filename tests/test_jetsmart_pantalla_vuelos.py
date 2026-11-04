"""Vuelos internacionales en el paso 3 de la pantalla de JetSmart: cantidad y
USD por vuelo por estacion, con la cuenta en vivo.

Correr con:
    python -m unittest tests.test_jetsmart_pantalla_vuelos -v

No necesita archivos reales ni la base (se quita DATABASE_URL).
"""

import os
import unittest
from pathlib import Path
from unittest import mock

from streamlit.testing.v1 import AppTest

APP = Path(__file__).resolve().parent.parent / "app.py"


def _app() -> AppTest:
    at = AppTest.from_file(str(APP), default_timeout=120)
    at.session_state["aerolinea_sel"] = "jetsmart"
    at.run()
    return at


def _textos(at: AppTest) -> str:
    return "\n".join([c.value for c in at.caption] + [w.value for w in at.warning])


class VuelosInternacionales(unittest.TestCase):
    def setUp(self):
        self._env = mock.patch.dict(os.environ)
        self._env.start()
        os.environ.pop("DATABASE_URL", None)
        self.at = _app()

    def tearDown(self):
        self._env.stop()

    def test_cuatro_campos_con_titulos_claros_y_defaults(self):
        self.assertEqual(self.at.number_input(key="js_vuelos_EZE").label, "Cantidad de vuelos")
        self.assertEqual(self.at.number_input(key="js_usd_EZE").label, "USD por vuelo")
        self.assertEqual(self.at.number_input(key="js_usd_EZE").value, 17.5)
        self.assertEqual(self.at.number_input(key="js_usd_MDZ").value, 35.0)
        self.assertEqual(self.at.number_input(key="js_vuelos_MDZ").value, 0)
        self.assertEqual(self.at.number_input(key="js_vuelos_EZE").label, self.at.number_input(key="js_vuelos_MDZ").label)

    def test_la_cantidad_es_entera_y_el_importe_acepta_decimales(self):
        cantidad, usd = self.at.number_input(key="js_vuelos_EZE"), self.at.number_input(key="js_usd_EZE")
        self.assertIsInstance(cantidad.value, int)
        self.assertEqual(cantidad.step, 1)
        self.assertEqual(usd.step, 0.5)
        self.assertIsInstance(usd.value, float)

    def test_cuenta_en_vivo_con_defaults(self):
        self.at.number_input(key="js_vuelos_EZE").set_value(49).run()
        self.at.number_input(key="js_vuelos_MDZ").set_value(38).run()
        textos = _textos(self.at)
        self.assertIn("49 vuelos × 17,5 USD = 857,5 USD", textos)
        self.assertIn("38 vuelos × 35 USD = 1.330 USD", textos)

    def test_la_cuenta_sigue_al_importe_editado(self):
        self.at.number_input(key="js_vuelos_EZE").set_value(49).run()
        self.at.number_input(key="js_usd_EZE").set_value(20.25).run()
        self.assertIn("49 vuelos × 20,25 USD = 992,25 USD", _textos(self.at))

    def test_un_solo_vuelo_va_en_singular(self):
        self.at.number_input(key="js_vuelos_EZE").set_value(1).run()
        self.assertIn("1 vuelo × 17,5 USD = 17,5 USD", _textos(self.at))


if __name__ == "__main__":
    unittest.main()
