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


def _eze_con_filas(expo: list, impo: list, tot_expo=None, tot_impo=None) -> list[list]:
    """Hoja de EZE con filas de vuelo reales. expo/impo: [(vuelo, mane, guias)]. Los totales tipeados
    salen de las filas (guias > 0 = con carga) salvo que se pasen (total, con_carga) a mano."""
    def totales(filas):
        con = sum(1 for _, _, g in filas if g)
        return len(filas), con, len(filas) - con
    te, ti = tot_expo or totales(expo)[:2], tot_impo or totales(impo)[:2]
    filas = [["Fecha", "Vuelo", "Mane", "Guias", "Destino", "Vuelo", "Mani", "Guias", "Origen"]]
    for k in range(max(len(expo), len(impo))):
        e = expo[k] if k < len(expo) else (None, None, None)
        i = impo[k] if k < len(impo) else (None, None, None)
        filas.append([None, e[0], e[1], e[2], None, i[0], i[1], i[2], None])
    filas += [[None, "TOTAL VUELOS EXPO", None, te[0], None, "TOTAL VUELOS IMPO", None, ti[0]]]
    filas += [[None, "CON CARGA", None, te[1], None, "CON CARGA", None, ti[1]]]
    filas += [[None, "SIN CARGA", None, te[0] - te[1], None, "SIN CARGA", None, ti[0] - ti[1]]]
    filas += [[], [None, 52.5]]
    return filas


class DesgloseYCeldaDelTotal(unittest.TestCase):
    def test_eze_devuelve_el_desglose_y_la_celda_del_total(self):
        filas = _eze_con_filas([("A1", "100 A", 0), ("A2", "101 B", 0)], [("B1", "200 C", 0)])
        eze = leer_vuelos_ariel(_libro({"mani mane EZE": filas}))["EZE"]
        self.assertEqual(eze.bloques, (("EXPO", 2), ("IMPO", 1)))
        self.assertEqual(eze.celda_total, "B8")  # encabezado + 2 filas + 3 de totales + vacia = 7; el total es la fila 8

    def test_mdz_devuelve_la_celda_del_total(self):
        filas = [["FECHA", "VUELO", "USD"], ["1/10", "JA1", 35], ["2/10", "JA2", 35], [None, None, 70]]
        mdz = leer_vuelos_ariel(_libro({"mani mane MDZ": filas}))["MDZ"]
        self.assertEqual((mdz.total_usd, mdz.celda_total), (70, "C4"))
        self.assertEqual(mdz.bloques, ())

    def test_sin_total_no_hay_celda(self):
        eze = leer_vuelos_ariel(_libro({"mani mane EZE": _eze(total=None)}))["EZE"]
        self.assertIsNone(eze.celda_total)


class ControlDeConteoEze(unittest.TestCase):
    """Se cuentan las filas de vuelo de Expo e Impo y se comparan con los totales tipeados."""

    def test_si_coinciden_no_avisa(self):
        filas = _eze_con_filas([("A1", "100 A", 0)] * 2 + [("A3", "102 C", 0)], [("B1", "200 A", 1), ("B2", "201 B", 0)])
        eze = leer_vuelos_ariel(_libro({"mani mane EZE": filas}))["EZE"]
        self.assertEqual(eze.avisos_conteo, ())

    def test_si_hay_mas_filas_que_el_total_tipeado_avisa_con_los_dos_numeros(self):
        filas = _eze_con_filas([("A1", "100 A", 0), ("A2", "101 B", 0), ("A3", "102 C", 0)], [("B1", "200 A", 0)],
                               tot_expo=(2, 0))
        eze = leer_vuelos_ariel(_libro({"mani mane EZE": filas}))["EZE"]
        self.assertEqual(eze.cantidad, 3)  # el aviso no cambia la cantidad
        self.assertEqual(len(eze.avisos_conteo), 1)
        self.assertIn("Exportación", eze.avisos_conteo[0])
        self.assertIn("conté 3 filas de vuelo", eze.avisos_conteo[0])
        self.assertIn("dicen 2 vuelos", eze.avisos_conteo[0])

    def test_si_las_con_carga_no_coinciden_avisa(self):
        filas = _eze_con_filas([("A1", "100 A", 0)], [("B1", "200 A", 1), ("B2", "201 B", 0)], tot_impo=(2, 0))
        eze = leer_vuelos_ariel(_libro({"mani mane EZE": filas}))["EZE"]
        self.assertEqual(len(eze.avisos_conteo), 1)
        self.assertIn("Importación", eze.avisos_conteo[0])
        self.assertIn("1 con carga", eze.avisos_conteo[0])

    def test_las_filas_sin_vuelos_no_se_cuentan(self):
        filas = _eze_con_filas([("A1", "100 A", 0), ("SIN VUELOS", None, None)], [("B1", "200 A", 0)],
                               tot_expo=(1, 0))
        eze = leer_vuelos_ariel(_libro({"mani mane EZE": filas}))["EZE"]
        self.assertEqual(eze.avisos_conteo, ())


class ManeRepetidoEze(unittest.TestCase):
    def test_avisa_si_el_mismo_mane_esta_en_dos_filas(self):
        filas = _eze_con_filas([("A1", "51116 D", 0), ("A2", "51116 D", 0), ("A3", "51117 E", 0)], [("B1", "200 A", 0)])
        eze = leer_vuelos_ariel(_libro({"mani mane EZE": filas}))["EZE"]
        self.assertEqual(eze.cantidad, 4)
        self.assertEqual(len(eze.avisos_repetidos), 1)
        self.assertIn("51116 D", eze.avisos_repetidos[0])
        self.assertIn("filas 2 y 3", eze.avisos_repetidos[0])
        self.assertIn("vuelo repetido", eze.avisos_repetidos[0])
        self.assertEqual(eze.avisos_conteo, ())

    def test_distinto_mane_no_avisa(self):
        filas = _eze_con_filas([("A1", "100 A", 0), ("A2", "101 B", 0)], [("B1", "100 A", 0)])  # mismo numero pero en el otro bloque
        eze = leer_vuelos_ariel(_libro({"mani mane EZE": filas}))["EZE"]
        self.assertEqual(eze.avisos_repetidos, ())


class CeldasDeLosImportesPorBloque(unittest.TestCase):
    """Las celdas "Vuelos Expo" / "Vuelos Impo" de la hoja de Anita, para citarlas en pantalla."""

    def _hoja(self, rotulos):
        filas = _eze_con_filas([("A1", "100 A", 0)], [("B1", "200 A", 0)])
        return filas[:-2] + [[], *rotulos, [], [None, 52.5]]

    def test_ubica_expo_e_impo_por_su_rotulo(self):
        filas = self._hoja([["Vuelos Expo", 17.5], ["Vuelos Impo", 17.5]])
        eze = leer_vuelos_ariel(_libro({"mani mane EZE": filas}))["EZE"]
        self.assertEqual(eze.celdas_bloques, (("EXPO", "B7"), ("IMPO", "B8")))

    def test_si_un_bloque_no_se_ubica_solo_falta_ese(self):
        filas = self._hoja([["Vuelos Expo", 17.5], ["Otra cosa", 3]])
        eze = leer_vuelos_ariel(_libro({"mani mane EZE": filas}))["EZE"]
        self.assertEqual(eze.celdas_bloques, (("EXPO", "B7"),))
        self.assertEqual(eze.cantidad, 2)  # no afecta la lectura

    def test_rotulo_repetido_no_se_adivina(self):
        filas = self._hoja([["Vuelos Expo", 17.5], ["Vuelos Expo", 20], ["Vuelos Impo", 17.5]])
        eze = leer_vuelos_ariel(_libro({"mani mane EZE": filas}))["EZE"]
        self.assertEqual(eze.celdas_bloques, (("IMPO", "B9"),))

    @unittest.skipUnless(ARIEL_OCT.exists(), "falta el Ariel de octubre")
    def test_octubre_expo_b39_e_impo_b40(self):
        self.assertEqual(leer_vuelos_ariel(ARIEL_OCT)["EZE"].celdas_bloques, (("EXPO", "B39"), ("IMPO", "B40")))


@unittest.skipUnless(ARIEL_OCT.exists(), "falta el Ariel de octubre")
class ControlesConArchivosReales(unittest.TestCase):
    def test_octubre_desglose_celdas_y_mane_repetido(self):
        v = leer_vuelos_ariel(ARIEL_OCT)
        self.assertEqual(v["EZE"].bloques, (("EXPO", 25), ("IMPO", 24)))
        self.assertEqual((v["EZE"].celda_total, v["MDZ"].celda_total), ("B42", "G30"))
        self.assertEqual(v["EZE"].avisos_conteo, ())
        self.assertEqual(len(v["EZE"].avisos_repetidos), 1)
        self.assertIn("51116 D", v["EZE"].avisos_repetidos[0])
        self.assertIn("filas 22 y 23", v["EZE"].avisos_repetidos[0])
        self.assertEqual(v["MDZ"].avisos_repetidos, ())

    @unittest.skipUnless(ARIEL_SEP.exists() and ARIEL_AGO.exists(), "faltan septiembre y agosto")
    def test_septiembre_y_agosto_no_dan_falsos_avisos(self):
        for ruta in (ARIEL_SEP, ARIEL_AGO):  # tienen filas "SIN VUELOS" que no son vuelos
            eze = leer_vuelos_ariel(ruta)["EZE"]
            self.assertEqual((eze.avisos_conteo, eze.avisos_repetidos), ((), ()), ruta.name)


if __name__ == "__main__":
    unittest.main()
