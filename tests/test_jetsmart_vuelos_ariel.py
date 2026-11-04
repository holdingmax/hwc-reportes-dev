"""Lectura de la cantidad de vuelos desde las hojas de manifiestos del Excel de Ariel.

Correr con:
    python -m unittest tests.test_jetsmart_vuelos_ariel -v

Los casos con archivos reales (octubre, agosto, septiembre 2026) usan
tests/fixtures/jetsmart/ (no versionado): si faltan, se saltean. Los casos
sinteticos no necesitan nada.
"""

import io
import unittest
from pathlib import Path

from openpyxl import Workbook

from src.jetsmart_vuelos_ariel import leer_vuelos_ariel

FIX = Path(__file__).resolve().parent / "fixtures" / "jetsmart"
ARIEL_OCT = FIX / "Ariel_Rendicion_octubre_2026.xlsx"
ARIEL_AGO = FIX / "Ariel_Rendicion_agosto_2026.xlsx"
ARIEL_SEP = FIX / "Ariel_Rendicion_septiembre_2026.xlsx"
GUIAS_OCT = FIX / "Guias_Aerea_10-2026.xlsx"


def _libro(hojas: dict[str, list[list]]) -> io.BytesIO:
    wb = Workbook()
    wb.remove(wb.active)
    for nombre, filas in hojas.items():
        ws = wb.create_sheet(nombre)
        for fila in filas:
            ws.append(fila)
    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer


def _eze(desplazar: int = 0, expo=(25, 0, 25), impo=(25, 1, 24), total=857.5) -> list[list]:
    """Hoja de EZE con los dos bloques al pie; `desplazar` filas vacias antes (los rotulos no estan en celda fija)."""
    filas = [[None] * 9 for _ in range(desplazar)]
    filas += [["Fecha", "Vuelo", "Mane", "Guias", "Destino", "Vuelo", "Mani", "Guias", "Origen"]]
    filas += [[None, "TOTAL VUELOS EXPO", None, expo[0], None, "TOTAL VUELOS IMPO", None, impo[0]]]
    filas += [[None, "CON CARGA", None, expo[1], None, "CON CARGA", None, impo[1]]]
    filas += [[None, "SIN CARGA", None, expo[2], None, "SIN CARGA", None, impo[2]]]
    if total is not None:
        filas += [[], [None, total]]
    return filas


@unittest.skipUnless(ARIEL_OCT.exists(), "falta el Ariel de octubre")
class ArielOctubre(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.vuelos = leer_vuelos_ariel(ARIEL_OCT)

    def test_eze_son_25_mas_24_sin_carga(self):
        eze = self.vuelos["EZE"]
        self.assertEqual(eze.cantidad, 49)
        self.assertIsNone(eze.problema)
        self.assertEqual(eze.origen, "Tomado del Excel de Ariel, hoja mani mane EZE: 25 + 24 vuelos sin carga = 49")

    def test_mdz_son_27_filas_de_vuelo_con_usd(self):
        mdz = self.vuelos["MDZ"]
        self.assertEqual(mdz.cantidad, 27)
        self.assertEqual(mdz.origen, "Tomado del Excel de Ariel, hoja mani mane MDZ: 27 vuelos con USD")

    def test_totales_en_usd_del_excel_para_controlar(self):
        self.assertEqual(self.vuelos["EZE"].total_usd, 857.5)  # 49 x 17,5
        self.assertEqual(self.vuelos["MDZ"].total_usd, 945)  # 27 x 35


@unittest.skipUnless(ARIEL_AGO.exists() and ARIEL_SEP.exists(), "faltan los Ariel de agosto y septiembre")
class OtrosMeses(unittest.TestCase):
    """Los nombres de hoja y los rotulos de MDZ cambian de un mes a otro."""

    def test_agosto(self):
        v = leer_vuelos_ariel(ARIEL_AGO)
        self.assertEqual((v["EZE"].cantidad, v["EZE"].total_usd), (26, 455))  # 14 + 12; rotulado "TOTAL"
        self.assertEqual((v["MDZ"].cantidad, v["MDZ"].total_usd), (36, 1260))  # columna "IMPORTE"

    def test_septiembre_sin_total_en_usd_en_eze(self):
        v = leer_vuelos_ariel(ARIEL_SEP)
        self.assertEqual(v["EZE"].cantidad, 16)  # 9 + 7
        self.assertIsNone(v["EZE"].total_usd)
        self.assertEqual((v["MDZ"].cantidad, v["MDZ"].total_usd), (33, 1155))


class SinHojasDeManifiestos(unittest.TestCase):
    def test_excel_sin_esas_hojas_no_rompe_y_no_trae_nada(self):
        self.assertEqual(leer_vuelos_ariel(_libro({"venta": [["a", "b"], [1, 2]]})), {})

    @unittest.skipUnless(GUIAS_OCT.exists(), "faltan las guias de octubre")
    def test_otro_archivo_real(self):
        self.assertEqual(leer_vuelos_ariel(GUIAS_OCT), {})

    def test_archivo_que_no_es_un_excel(self):
        self.assertEqual(leer_vuelos_ariel(io.BytesIO(b"esto no es un xlsx")), {})

    def test_solo_una_estacion(self):
        v = leer_vuelos_ariel(_libro({"mani mane EZE": _eze()}))
        self.assertEqual(list(v), ["EZE"])
        self.assertEqual(v["EZE"].cantidad, 49)


class SinCertezaNoCompleta(unittest.TestCase):
    def test_los_rotulos_se_buscan_no_estan_en_celda_fija(self):
        for filas_antes in (0, 7, 30):
            with self.subTest(filas_antes=filas_antes):
                self.assertEqual(leer_vuelos_ariel(_libro({"mani mane EZE": _eze(filas_antes)}))["EZE"].cantidad, 49)

    def test_un_solo_bloque_no_se_completa(self):
        filas = [r[:5] for r in _eze()]  # sin el bloque IMPO
        eze = leer_vuelos_ariel(_libro({"mani mane EZE": filas}))["EZE"]
        self.assertIsNone(eze.cantidad)
        self.assertIn("SIN CARGA", eze.problema)

    def test_bloque_que_no_cierra_no_se_completa(self):
        eze = leer_vuelos_ariel(_libro({"mani mane EZE": _eze(impo=(25, 1, 20))}))["EZE"]  # 1 + 20 != 25
        self.assertIsNone(eze.cantidad)
        self.assertIn("no cierra", eze.problema)

    def test_valor_que_no_es_numero_no_se_completa(self):
        eze = leer_vuelos_ariel(_libro({"mani mane EZE": _eze(expo=(25, 0, "veinticinco"))}))["EZE"]
        self.assertIsNone(eze.cantidad)
        self.assertIsNotNone(eze.problema)

    def test_mdz_sin_columna_de_vuelos(self):
        mdz = leer_vuelos_ariel(_libro({"mani mane MDZ": [["FECHA", "OBSERVACIONES"], ["1/10", "x"]]}))["MDZ"]
        self.assertIsNone(mdz.cantidad)
        self.assertIn("VUELO", mdz.problema)

    def test_mdz_ignora_el_total_y_cuenta_solo_vuelos_con_importe(self):
        filas = [["FECHA", "VUELO", "USD"], ["1/10", "JA1", 35], ["2/10", "JA2", 35], ["3/10", "JA3", None], [None, None, 70]]
        mdz = leer_vuelos_ariel(_libro({"mani mane MDZ": filas}))["MDZ"]
        self.assertEqual((mdz.cantidad, mdz.total_usd), (2, 70))

    def test_dos_hojas_de_la_misma_estacion_no_se_completa(self):
        v = leer_vuelos_ariel(_libro({"mani mane EZE": _eze(), "mane EZE vieja": _eze()}))
        self.assertIsNone(v["EZE"].cantidad)
        self.assertIn("más de una hoja", v["EZE"].problema)


if __name__ == "__main__":
    unittest.main()
