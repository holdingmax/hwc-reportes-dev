"""Regresion del cierre de octubre 2026 con los archivos reales: guias del
sistema de octubre + rendicion de Ariel de octubre (sin repetidas), con
septiembre como mes anterior.

Resultado validado por Anita (4/11/2026) con un mapeo provisorio, y
reproducido aca con el mapeo definitivo de src/jetsmart_archivos.py:
- 1.953 guias incluidas = 1.934 que matchean + 19 "pendiente del mes
  anterior, declarada ahora" (vuelos del 1 y 2/10 creados en septiembre);
- 10 pendientes de declarar (creadas el 30 y 31/10, vuelan en noviembre);
- 0 diferencias de kg o de ingreso.

Correr con:
    python -m unittest tests.test_jetsmart_regresion_octubre -v

Usa tests/fixtures/jetsmart/ (no versionado): si faltan, se saltea.
"""

import unittest
from pathlib import Path

import pandas as pd

from src.config import COL_JS_GUIA, COL_JS_KGS, COL_JS_PRIORIDAD
from src.jetsmart_archivos import cargar_guias_sistema
from src.jetsmart_builder import build_jetsmart_guias
from src.jetsmart_cruce import (
    EXC_DIFERENCIA,
    EXC_DOBLE,
    EXC_DUPLICADA,
    EXC_PENDIENTE,
    EXC_REVISION,
    INC_MATCH,
    INC_PENDIENTE_ANTERIOR,
    INC_ULTIMO_DIA,
    cruzar_jetsmart,
    leer_ariel,
)

FIX = Path(__file__).resolve().parent / "fixtures" / "jetsmart"
ARCHIVOS = [
    FIX / "Guias_Aerea_10-2026.xlsx", FIX / "Guias_Aerea_09-2026.xlsx",
    FIX / "Ariel_Rendicion_octubre_2026.xlsx", FIX / "Ariel_Rendicion_septiembre_2026.xlsx",
]


@unittest.skipUnless(all(p.exists() for p in ARCHIVOS), "faltan los archivos de septiembre/octubre en tests/fixtures/jetsmart/")
class CierreOctubre2026(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.guias_oct = cargar_guias_sistema(ARCHIVOS[0])
        cls.guias_sep = cargar_guias_sistema(ARCHIVOS[1])
        cls.ariel_oct = leer_ariel(ARCHIVOS[2], cruda=True)
        cls.ariel_sep = leer_ariel(ARCHIVOS[3], cruda=True)
        cls.r = cruzar_jetsmart(
            cls.guias_oct, cls.ariel_oct.guias, ("OCT", "26"), cls.guias_sep, cls.ariel_sep.guias,
            ariel_repetidas_identicas=cls.ariel_oct.repetidas_identicas,
        )

    def test_incluidas(self):
        self.assertEqual(len(self.r.incluidas), 1953)
        self.assertEqual(self.r.resumen["motivos_inclusion"], {INC_MATCH: 1934, INC_PENDIENTE_ANTERIOR: 19})

    def test_son_exactamente_las_guias_de_ariel_sin_repetidas(self):
        # Ariel es la fuente de la liquidacion: entra todo lo que declaro, una vez cada guia.
        self.assertEqual(set(self.r.incluidas[COL_JS_GUIA]), set(self.ariel_oct.guias["Nro de guía"]))
        self.assertFalse(self.r.incluidas[COL_JS_GUIA].duplicated().any())

    def test_diez_pendientes_de_declarar_y_nada_mas(self):
        self.assertEqual(self.r.resumen["excepciones_por_tipo"], {EXC_PENDIENTE: 10})
        self.assertEqual(len(self.r.excepciones), 10)
        # Creadas el 30 y 31/10, todavia sin declarar: salen en noviembre.
        pendientes = set(self.r.excepciones["Nro guía"])
        creadas = self.guias_oct[pd.to_numeric(self.guias_oct[COL_JS_GUIA]).isin(pendientes)]["Creacion"]
        self.assertEqual(len(creadas), 10)
        self.assertEqual(set(creadas), {"30/10/2026", "31/10/2026"})

    def test_cero_diferencias_de_kg_o_ingreso(self):
        self.assertNotIn(EXC_DIFERENCIA, set(self.r.excepciones["Tipo"]))
        inc = self.r.incluidas.set_index(COL_JS_GUIA)
        base = self.guias_oct.assign(**{COL_JS_GUIA: pd.to_numeric(self.guias_oct[COL_JS_GUIA])}).set_index(COL_JS_GUIA)
        comunes = inc.index.intersection(base.index)
        self.assertEqual(len(comunes), 1934)
        self.assertTrue((inc.loc[comunes, COL_JS_KGS].astype(float) == base.loc[comunes, COL_JS_KGS].astype(float)).all())
        self.assertTrue((inc.loc[comunes, COL_JS_PRIORIDAD].astype(float) == base.loc[comunes, COL_JS_PRIORIDAD].astype(float)).all())

    def test_sin_otras_excepciones_ni_avisos_de_ultimo_dia(self):
        for tipo in (EXC_DOBLE, EXC_DUPLICADA, EXC_REVISION):
            self.assertNotIn(tipo, set(self.r.excepciones["Tipo"]))
        self.assertEqual(len(self.r.incluidas_ultimo_dia), 0)
        self.assertNotIn(INC_ULTIMO_DIA, self.r.resumen["motivos_inclusion"])
        self.assertEqual(self.r.resumen["ya_declaradas_mes_anterior"], 0)

    def test_las_20_repetidas_se_cuentan_una_vez_y_se_avisan(self):
        self.assertEqual(len(self.r.resumen["ariel_repetidas_identicas"]), 20)

    def test_las_guias_armadas_para_la_liquidacion(self):
        guias = build_jetsmart_guias(self.r.incluidas)
        self.assertEqual(len(guias), 1953)
        self.assertAlmostEqual(float(guias["GRAVADO"].sum()), float(self.ariel_oct.guias["Ingreso total flete"].sum()), places=2)


if __name__ == "__main__":
    unittest.main()
