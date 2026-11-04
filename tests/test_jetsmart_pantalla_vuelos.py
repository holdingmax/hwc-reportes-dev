"""Vuelos internacionales en el paso 3 de la pantalla de JetSmart: cantidad y
USD por vuelo por estacion, con la cuenta en vivo.

Correr con:
    python -m unittest tests.test_jetsmart_pantalla_vuelos -v

No necesita archivos reales ni la base (se quita DATABASE_URL).
"""

import io
import os
import unittest
from pathlib import Path
from unittest import mock

from openpyxl import load_workbook
from streamlit.testing.v1 import AppTest

FIX = Path(__file__).resolve().parent / "fixtures" / "jetsmart"
ARIEL_OCT = FIX / "Ariel_Rendicion_octubre_2026.xlsx"
ARIEL_SEP = FIX / "Ariel_Rendicion_septiembre_2026.xlsx"
GUIAS_OCT = FIX / "Guias_Aerea_10-2026.xlsx"
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
APP = Path(__file__).resolve().parent.parent / "app.py"


def _app() -> AppTest:
    at = AppTest.from_file(str(APP), default_timeout=120)
    at.session_state["aerolinea_sel"] = "jetsmart"
    at.run()
    return at


def _textos(at: AppTest) -> str:
    return "\n".join([c.value for c in at.caption] + [w.value for w in at.warning])


def _subir(at: AppTest, *archivos: tuple[str, bytes]) -> AppTest:
    at.get("file_uploader")[0].set_value([(nombre, datos, XLSX) for nombre, datos in archivos])
    return at.run()


def _ariel_sin_manifiestos() -> bytes:
    wb = load_workbook(ARIEL_OCT)
    for hoja in wb.sheetnames:
        if hoja != "venta":
            del wb[hoja]
    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


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


    def test_aviso_amarillo_si_la_cantidad_supera_150_por_estacion(self):
        self.at.number_input(key="js_vuelos_EZE").set_value(151).run()
        avisos = [w.value for w in self.at.warning]
        self.assertEqual(avisos.count("¿Es la cantidad o el importe?"), 1)
        # La cuenta se sigue mostrando: el aviso no bloquea nada.
        self.assertIn("151 vuelos", _textos(self.at))
        self.assertEqual(len(self.at.exception), 0)

    def test_sin_aviso_hasta_150_inclusive(self):
        self.at.number_input(key="js_vuelos_EZE").set_value(150).run()
        self.at.number_input(key="js_vuelos_MDZ").set_value(49).run()
        self.assertEqual([w.value for w in self.at.warning], [])

    def test_el_aviso_es_por_estacion(self):
        self.at.number_input(key="js_vuelos_MDZ").set_value(1000).run()
        self.assertEqual([w.value for w in self.at.warning], ["¿Es la cantidad o el importe?"])
        self.at.number_input(key="js_vuelos_EZE").set_value(200).run()
        self.assertEqual(len(self.at.warning), 2)


@unittest.skipUnless(ARIEL_OCT.exists() and ARIEL_SEP.exists() and GUIAS_OCT.exists(), "faltan los archivos reales")
class VuelosDesdeElExcelDeAriel(unittest.TestCase):
    def setUp(self):
        self._env = mock.patch.dict(os.environ)
        self._env.start()
        os.environ.pop("DATABASE_URL", None)
        self.at = _app()

    def tearDown(self):
        self._env.stop()

    def test_precarga_las_cantidades_y_dice_de_donde_salen(self):
        _subir(self.at, (ARIEL_OCT.name, ARIEL_OCT.read_bytes()))
        self.assertEqual(self.at.number_input(key="js_vuelos_EZE").value, 49)
        self.assertEqual(self.at.number_input(key="js_vuelos_MDZ").value, 27)
        textos = _textos(self.at)
        self.assertIn("Tomado del Excel de Ariel, hoja mani mane EZE: 25 + 24 vuelos sin carga = 49", textos)
        self.assertIn("Tomado del Excel de Ariel, hoja mani mane MDZ: 27 vuelos con USD", textos)
        self.assertIn("49 vuelos × 17,5 USD = 857,5 USD", textos)
        self.assertEqual(len(self.at.exception), 0)

    def test_lo_corregido_a_mano_no_se_pisa_y_se_avisa(self):
        _subir(self.at, (ARIEL_OCT.name, ARIEL_OCT.read_bytes()))
        self.at.number_input(key="js_vuelos_EZE").set_value(48).run()
        self.assertEqual(self.at.number_input(key="js_vuelos_EZE").value, 48)
        textos = _textos(self.at)
        self.assertIn("Cambiaste la cantidad a mano: el Excel de Ariel (hoja mani mane EZE) tenía 49", textos)
        self.assertNotIn("Tomado del Excel de Ariel, hoja mani mane EZE", textos)
        self.at.run()  # otra interaccion cualquiera: sigue en 48
        self.assertEqual(self.at.number_input(key="js_vuelos_EZE").value, 48)

    def test_otro_archivo_de_ariel_vuelve_a_precargar(self):
        _subir(self.at, (ARIEL_OCT.name, ARIEL_OCT.read_bytes()))
        self.at.number_input(key="js_vuelos_EZE").set_value(48).run()
        _subir(self.at, (ARIEL_SEP.name, ARIEL_SEP.read_bytes()))
        self.assertEqual(self.at.number_input(key="js_vuelos_EZE").value, 16)
        self.assertEqual(self.at.number_input(key="js_vuelos_MDZ").value, 33)

    def test_ariel_sin_las_hojas_de_manifiestos_no_rompe_y_deja_los_campos_vacios(self):
        _subir(self.at, ("Ariel sin manifiestos.xlsx", _ariel_sin_manifiestos()))
        self.assertEqual(len(self.at.exception), 0)
        self.assertEqual(self.at.number_input(key="js_vuelos_EZE").value, 0)
        self.assertEqual(self.at.number_input(key="js_vuelos_MDZ").value, 0)
        self.assertNotIn("Tomado del Excel de Ariel", _textos(self.at))

    def test_sin_excel_de_ariel_los_campos_siguen_a_mano(self):
        _subir(self.at, (GUIAS_OCT.name, GUIAS_OCT.read_bytes()))
        self.assertEqual(len(self.at.exception), 0)
        self.at.number_input(key="js_vuelos_EZE").set_value(12).run()
        self.assertEqual(self.at.number_input(key="js_vuelos_EZE").value, 12)
        self.assertNotIn("Excel de Ariel", _textos(self.at))


if __name__ == "__main__":
    unittest.main()
