"""Regresion del cruce JetSmart contra el cruce MANUAL real de Anita: la
hoja CONTROL de su archivo de trabajo de agosto 2026.

Correr con:
    python -m unittest tests.test_jetsmart_control_agosto -v

Necesita en tests/fixtures/jetsmart/ (no versionados, datos del cliente):
PRUEBA_LIQUIDACION_08-2026.xlsx (BD + ARIEL + CONTROL de agosto) y
LIQUIDACION_07-2026.xlsx (julio, mes anterior).

La hoja CONTROL pone lado a lado Ariel (columnas A-L) y el export (P-U),
alineados a mano, con 3 restas: control 1 = guia Ariel - guia export,
control 2 = Kg, control 3 = ingreso. En agosto esos valores vienen
pegados como numeros (sin formula). Lo que marca Anita:
- control 1 = el propio numero de guia (positivo): la guia esta en Ariel
  y no en el export. Son 6 guias del 05/08, vuelo 3232 AEP-TUC
  (96816 ... 96847), que al final de la hoja vuelve a listar como
  "ANULADAS": su resolucion fue no liquidarlas.
- control 1 negativo (-98164, -98186, -98191): al reves, estaban en su
  export y no en Ariel; les anoto "PASAR A SEPT". Ariel las declara en
  septiembre.
- control 1 = "ANULADAS" (texto): ese listado de resolucion, no una
  diferencia nueva; no se interpreta aca.

No se busca un match fila por fila: se anclan estos subconjuntos conocidos
para que un cambio futuro en el cruce que los rompa salte aca.
"""

import unittest
from pathlib import Path

import pandas as pd

from src.jetsmart_builder import detect_jetsmart_period, load_jetsmart_export
from src.jetsmart_cruce import EXC_REVISION, cruzar_jetsmart, load_ariel

FIX = Path(__file__).resolve().parent / "fixtures" / "jetsmart"
TRABAJO_AGO = FIX / "PRUEBA_LIQUIDACION_08-2026.xlsx"
TRABAJO_JUL = FIX / "LIQUIDACION_07-2026.xlsx"

SEIS_DEL_05_08 = {96816, 96823, 96829, 96833, 96839, 96847}
PASAR_A_SEPT = {98164, 98186, 98191}


def _control_anita() -> pd.DataFrame:
    with pd.ExcelFile(TRABAJO_AGO) as xls:
        hoja = next(s for s in xls.sheet_names if s.strip().upper() == "CONTROL")
        df = pd.read_excel(xls, sheet_name=hoja)
    df["_control1"] = pd.to_numeric(df["control 1"], errors="coerce")  # "ANULADAS" -> NaN
    return df


@unittest.skipUnless(TRABAJO_AGO.exists() and TRABAJO_JUL.exists(), "faltan los archivos de tests/fixtures/jetsmart/")
class TestCruceContraControlDeAnita(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.control = _control_anita()
        export_ago = load_jetsmart_export(TRABAJO_AGO)
        ariel_ago, _ = load_ariel(TRABAJO_AGO)
        export_jul = load_jetsmart_export(TRABAJO_JUL)
        ariel_jul, _ = load_ariel(TRABAJO_JUL)
        periodo, _ = detect_jetsmart_period(export_ago)
        # mismo armado que la validacion manual: agosto contra julio, sin mes siguiente
        cls.cruce = cruzar_jetsmart(export_ago, ariel_ago, periodo, export_jul, ariel_jul)
        cls.excepciones = set(cls.cruce.excepciones["Nro guía"])
        cls.incluidas = set(cls.cruce.incluidas["# Guía"])

    def test_control_1_distinto_de_cero_es_el_subconjunto_conocido(self):
        marcadas = self.control.loc[
            self.control["_control1"].notna() & (self.control["_control1"] != 0) & self.control["Nro de guía"].notna(),
            "Nro de guía",
        ]
        self.assertEqual(set(marcadas.astype(int)), SEIS_DEL_05_08)

    def test_las_marcadas_por_anita_son_excepcion_del_cruce(self):
        self.assertLessEqual(SEIS_DEL_05_08, self.excepciones)
        exc = self.cruce.excepciones
        tipos = set(exc.loc[exc["Nro guía"].isin(SEIS_DEL_05_08), "Tipo"])
        self.assertEqual(tipos, {EXC_REVISION})
        # y, como resolvio Anita ("ANULADAS"), no entran a la liquidacion
        self.assertFalse(SEIS_DEL_05_08 & self.incluidas)

    def test_pasar_a_septiembre_no_se_liquida_en_agosto(self):
        negativas = self.control.loc[self.control["_control1"] < 0, "_control1"]
        self.assertEqual(set((-negativas).astype(int)), PASAR_A_SEPT)
        self.assertFalse(PASAR_A_SEPT & self.incluidas)


if __name__ == "__main__":
    unittest.main()
