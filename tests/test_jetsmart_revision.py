"""Guias para revisar de JetSmart, explicadas en palabras simples (que paso y que hacer).

Correr con:
    python -m unittest tests.test_jetsmart_revision -v
"""

import unittest

import pandas as pd

from src.jetsmart_cruce import EXC_DIFERENCIA, EXC_DOBLE, EXC_DUPLICADA, EXC_PENDIENTE, EXC_REVISION
from src.jetsmart_revision import filas_para_revisar, resumen_por_tipo

COLUMNAS = ["Tipo", "Nro guía", "Fecha Ariel", "Origen", "Kg export", "Kg Ariel", "Ingreso export",
            "Ingreso Ariel", "¿Entra a la liquidación?", "Detalle"]


def _exc(*filas) -> pd.DataFrame:
    """filas: (tipo, nro guia, entra, detalle)."""
    return pd.DataFrame(
        [[t, g, None, "AEP", None, None, None, None, e, d] for t, g, e, d in filas], columns=COLUMNAS,
    )


class FilasParaRevisar(unittest.TestCase):
    def test_sin_excepciones_no_hay_filas(self):
        self.assertEqual(filas_para_revisar(_exc()), [])
        self.assertEqual(resumen_por_tipo(_exc()), [])

    def test_cada_tipo_dice_que_paso_y_que_hacer(self):
        filas = filas_para_revisar(_exc(
            (EXC_PENDIENTE, 1, "No", "x"), (EXC_REVISION, 2, "No", "x"), (EXC_DUPLICADA, 3, "No", "x"),
            (EXC_DIFERENCIA, 4, "Sí, con valores de Ariel", "Difiere Kg y ingreso."), (EXC_DOBLE, 5, "No", "x"),
        ))
        self.assertEqual([f["guia"] for f in filas], [1, 2, 3, 4, 5])
        for f in filas:
            self.assertTrue(f["que_pasa"] and f["que_hacer"])
        pendiente = filas[0]
        self.assertIn("Ariel todavía no la declaró", pendiente["que_pasa"])
        self.assertIn("Esperá a que Ariel la declare", pendiente["que_hacer"])
        self.assertEqual(pendiente["entra"], "No")

    def test_diferencia_dice_en_que_difiere_y_que_entra_con_valores_de_ariel(self):
        (f,) = filas_para_revisar(_exc((EXC_DIFERENCIA, 9, "Sí, con valores de Ariel", "Difiere Kg y ingreso.")))
        self.assertIn("Kg y ingreso", f["que_pasa"])
        self.assertEqual(f["entra"], "Sí (valores de Ariel)")
        self.assertIn("Ya entra con los valores de Ariel", f["que_hacer"])

    def test_numero_de_guia_entero_aunque_llegue_decimal(self):
        filas = filas_para_revisar(_exc((EXC_PENDIENTE, 102176.99999999999, "No", "x"), (EXC_PENDIENTE, 102177.0, "No", "x")))
        self.assertEqual([f["guia"] for f in filas], [102177, 102177])
        self.assertTrue(all(isinstance(f["guia"], int) for f in filas))

    def test_los_pendientes_van_primero_y_por_numero_de_guia(self):
        filas = filas_para_revisar(_exc((EXC_DOBLE, 1, "No", "x"), (EXC_PENDIENTE, 30, "No", "x"), (EXC_PENDIENTE, 20, "No", "x")))
        self.assertEqual([f["guia"] for f in filas], [20, 30, 1])

    def test_un_tipo_desconocido_no_se_esconde(self):
        (f,) = filas_para_revisar(_exc(("Tipo nuevo", 7, "No", "detalle tecnico")))
        self.assertEqual(f["guia"], 7)
        self.assertIn("detalle tecnico", f["que_pasa"])
        self.assertEqual(resumen_por_tipo(_exc(("Tipo nuevo", 7, "No", "d"))), ["1 otras"])

    def test_resumen_con_singular_y_plural(self):
        resumen = resumen_por_tipo(_exc(
            (EXC_PENDIENTE, 1, "No", ""), (EXC_PENDIENTE, 2, "No", ""), (EXC_DIFERENCIA, 3, "Sí, con valores de Ariel", "Difiere Kg."),
        ))
        self.assertEqual(resumen, ["2 pendientes de declarar", "1 con diferencias"])


if __name__ == "__main__":
    unittest.main()
