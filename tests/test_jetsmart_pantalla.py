"""Pantalla de JetSmart en la app real (zona unica de carga), manejada con
el framework de tests de Streamlit.

Correr con:
    python -m unittest tests.test_jetsmart_pantalla -v

Usa los archivos reales de octubre y septiembre 2026 en
tests/fixtures/jetsmart/ (no versionado): si faltan, se saltea. Nunca
escribe en la base: se quita DATABASE_URL del entorno y se reemplaza el
guardado por un doble.
"""

import os
import unittest
from pathlib import Path
from unittest import mock

from streamlit.testing.v1 import AppTest

FIX = Path(__file__).resolve().parent / "fixtures" / "jetsmart"
ARCHIVOS = {
    "guias_oct": FIX / "Guias_Aerea_10-2026.xlsx",
    "listado_oct": FIX / "Listado_Facturacion_JM_10-2026.xlsx",
    "ariel_oct": FIX / "Ariel_Rendicion_octubre_2026.xlsx",
    "iva_oct": FIX / "IVA_VENTAS_HW_10-2026.xlsx",
    "ariel_sep": FIX / "Ariel_Rendicion_septiembre_2026.xlsx",
    "bd_sep": FIX / "PRUEBA_LIQUIDACION_09-2026.xlsx",   # listado de septiembre (hoja BD)
    "liq_armada": FIX / "LIQ_ECS_07-2026_00000005.xlsx",
}
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
APP = Path(__file__).resolve().parent.parent / "app.py"


def _app(claves: list[str]) -> AppTest:
    at = AppTest.from_file(str(APP), default_timeout=180)
    at.session_state["aerolinea_sel"] = "jetsmart"
    at.run()
    at.get("file_uploader")[0].set_value([(ARCHIVOS[k].name, ARCHIVOS[k].read_bytes(), XLSX) for k in claves])
    at.run()
    return at


def _textos(at: AppTest) -> str:
    partes = [m.value for m in at.markdown] + [w.value for w in at.warning] + [e.value for e in at.error]
    return "\n".join(partes)


@unittest.skipUnless(all(p.exists() for p in ARCHIVOS.values()), "faltan los archivos reales")
class PantallaJetSmart(unittest.TestCase):
    def setUp(self):
        # Sin base: ni lectura ni escritura (el guardado es fail-soft, pero aca ni se intenta).
        self._env = mock.patch.dict(os.environ)
        self._env.start()
        os.environ.pop("DATABASE_URL", None)
        self._guardar = mock.patch("src.db.guardar_liquidacion_jetsmart")
        self.guardar = self._guardar.start()

    def tearDown(self):
        self._guardar.stop()
        self._env.stop()

    def test_una_sola_zona_de_carga_con_varios_archivos(self):
        at = _app([])
        subidas = at.get("file_uploader")
        self.assertEqual(len(subidas), 1)
        self.assertTrue(subidas[0].accept_multiple_files)
        self.assertEqual(len(at.exception), 0)

    def test_textos_sin_nombres_internos(self):
        at = _app([])
        texto = _textos(at)
        self.assertIn("guías del mes", texto)
        self.assertIn("rendición de Ariel", texto)
        for interno in ("LIQUIDACION ECS", "hoja BD", "hoja ARIEL", "ValueError"):
            self.assertNotIn(interno, texto)

    def test_reconoce_cada_archivo_y_dice_que_es(self):
        at = _app(["iva_oct", "ariel_oct", "listado_oct", "guias_oct"])   # a proposito, en desorden
        texto = _textos(at)
        self.assertIn("Detecté: Guías del sistema de octubre 2026, 1.944 guías", texto)
        self.assertIn("Detecté: Rendición de Ariel de octubre 2026, 1.953 guías", texto)
        self.assertIn("Detecté: Listado de facturación de octubre 2026, 1.905 guías", texto)
        self.assertIn("Detecté: Libro IVA Ventas de AFIP de octubre 2026", texto)
        self.assertEqual(len(at.exception), 0)

    def test_sin_mes_anterior_explica_que_falta_y_no_deja_procesar(self):
        at = _app(["guias_oct", "ariel_oct"])
        texto = _textos(at)
        self.assertIn("Falta el mes anterior", texto)
        self.assertIn("septiembre 2026", texto)
        self.assertTrue(at.button[0].disabled)

    def test_archivo_equivocado_da_un_mensaje_de_negocio(self):
        at = _app(["guias_oct", "ariel_oct", "liq_armada"])
        errores = " ".join(e.value for e in at.error)
        self.assertIn("liquidación ya armada", errores)
        self.assertNotIn("ValueError", errores)
        self.assertEqual(len(at.exception), 0)

    def test_flujo_completo_genera_la_liquidacion(self):
        at = _app(["iva_oct", "ariel_sep", "guias_oct", "bd_sep", "ariel_oct"])
        texto = _textos(at)
        self.assertNotIn("Falta", texto)
        # Tipo de cambio: sin el, el boton sigue deshabilitado.
        self.assertTrue(at.button[0].disabled)
        at.number_input[0].set_value(1485.0)
        at.run()
        self.assertFalse(at.button[0].disabled)
        at.button[0].click()
        at.run()
        self.assertEqual(len(at.exception), 0)
        self.assertEqual([e.value for e in at.error], [])
        self.assertEqual(len(at.get("download_button")), 1)
        # Las 20 filas repetidas identicas de Ariel se avisan, en el resultado y en el resumen guardado.
        avisos = " ".join(w.value for w in at.warning)
        self.assertIn("20 filas repetidas idénticas", avisos)
        resumen_guardado = self.guardar.call_args.kwargs["cruce"]["resumen"]
        self.assertEqual(len(resumen_guardado["ariel_repetidas_identicas"]), 20)
        # Se guardo con las guias del sistema de octubre, no con el listado.
        kwargs = self.guardar.call_args.kwargs
        self.assertEqual(kwargs["nombre_archivo"], ARCHIVOS["guias_oct"].name)
        self.assertEqual(kwargs["period"], ("OCT", "26"))


if __name__ == "__main__":
    unittest.main()
