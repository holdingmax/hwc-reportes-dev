"""Garantia: las versiones de origen 'liquidacion' SIEMPRE guardan el
resultado del cruce en cada guia. Solo la carga inicial puede dejarlo vacio.

Correr con:
    python -m unittest tests.test_jetsmart_resultado_liquidacion -v

Por que importa: el mes siguiente tiene que saber que paso con TODAS las
guias del mes (incluidas, pendientes de declarar y excluidas). La columna
resultado de la tabla acepta NULL solo por la carga inicial (septiembre 2026,
liquidado fuera del sistema); esa excepcion no puede colarse en una
liquidacion. No se puede expresar como CHECK de la tabla (el origen vive en
otra tabla), asi que se valida en el codigo y se fija aca, con los meses
reales (agosto, septiembre y octubre 2026) y con casos armados a mano.
"""

import unittest
from pathlib import Path

import pandas as pd

from src.jetsmart_archivos import cargar_guias_sistema
from src.jetsmart_builder import load_jetsmart_export
from src.jetsmart_cruce import (
    RES_EXCLUIDA,
    RES_INCLUIDA,
    RES_PENDIENTE,
    cruzar_jetsmart,
    leer_ariel,
    load_ariel,
)
from src.jetsmart_mes_anterior import FUENTE_ARIEL, FUENTE_SISTEMA, filas_guias_periodo, validar_resultados
from tests.test_jetsmart_mes_anterior import (
    RESPUESTAS_GUARDAR,
    DestinosPorRegla,
    _Conexion,
    _guardar,
    _mes,
)

FIX = Path(__file__).resolve().parent / "fixtures" / "jetsmart"
RESULTADOS_VALIDOS = {RES_INCLUIDA, RES_PENDIENTE, RES_EXCLUIDA}


def _verificar(test: unittest.TestCase, export, ariel, destinos):
    """Arma las filas como lo hace la app para una liquidacion y comprueba que
    cada guia de las dos fuentes quedo con un resultado valido, igual en sus dos filas."""
    filas = filas_guias_periodo(export, ariel, destinos)   # lanza si falta alguno
    test.assertTrue(filas)
    test.assertTrue(all(f["resultado"] in RESULTADOS_VALIDOS for f in filas))
    por_guia = {}
    for f in filas:
        por_guia.setdefault(f["nro_guia"], set()).add((f["resultado"], f["motivo"]))
    test.assertTrue(all(len(v) == 1 for v in por_guia.values()), "las dos filas de una guia deben decir lo mismo")
    return filas


class ValidacionDeResultados(unittest.TestCase):
    def test_liquidacion_con_una_guia_sin_resultado_se_rechaza(self):
        filas = [
            {"fuente": FUENTE_SISTEMA, "nro_guia": 1, "resultado": RES_INCLUIDA},
            {"fuente": FUENTE_ARIEL, "nro_guia": 2, "resultado": None},
        ]
        with self.assertRaisesRegex(ValueError, "sin resultado"):
            validar_resultados(filas)

    def test_resultado_vacio_en_texto_tambien_se_rechaza(self):
        with self.assertRaises(ValueError):
            validar_resultados([{"fuente": FUENTE_ARIEL, "nro_guia": 2, "resultado": ""}])

    def test_el_armado_de_una_liquidacion_exige_destino_para_cada_guia(self):
        # Si el cruce dejara una guia sin destino, no se puede armar la version de liquidacion.
        from tests.test_jetsmart_duplicados import _export_con
        from src.jetsmart_cruce import _AR_COLUMNS
        fila = [827, 1, 5, pd.Timestamp(2026, 10, 15), 3039, "C", "D", "AEP", "BRC", 1, 8, 8, 736, 10000, "X"]
        ariel = pd.DataFrame([fila], columns=_AR_COLUMNS).assign(**{"Nro de guía": 5})
        with self.assertRaises(ValueError):
            filas_guias_periodo(_export_con(5), ariel, destinos={})          # sin destino para la guia 5
        self.assertTrue(filas_guias_periodo(_export_con(5), ariel, destinos={5: (RES_INCLUIDA, "match")}))

    def test_la_carga_inicial_si_puede_quedar_sin_resultado(self):
        from tests.test_jetsmart_duplicados import _export_con
        from src.jetsmart_cruce import _AR_COLUMNS
        fila = [827, 1, 5, pd.Timestamp(2026, 9, 15), 3039, "C", "D", "AEP", "BRC", 1, 8, 8, 736, 10000, "X"]
        ariel = pd.DataFrame([fila], columns=_AR_COLUMNS).assign(**{"Nro de guía": 5})
        filas = filas_guias_periodo(_export_con(5), ariel, destinos=None)
        self.assertTrue(all(f["resultado"] is None for f in filas))


class GuardadoEnLaBaseExigeResultado(unittest.TestCase):
    def test_una_liquidacion_sin_resultado_no_se_escribe(self):
        conn = _Conexion(RESPUESTAS_GUARDAR)
        r = _guardar(conn, [_mes("liquidacion", sin_resultado=True)])
        # La liquidacion en si se guarda (como siempre); las guias del mes, no.
        self.assertEqual(r, {"liquidacion": True, "meses": False})
        self.assertFalse([e for e in conn.log if e[0] == "executemany"])
        self.assertFalse([e for e in conn.log if e[0] == "execute" and "INSERT INTO jetsmart_periodos" in e[1]])

    def test_la_carga_inicial_sin_resultado_si_se_escribe(self):
        conn = _Conexion(RESPUESTAS_GUARDAR)
        r = _guardar(conn, [_mes("carga_inicial", 3)])
        self.assertEqual(r, {"liquidacion": True, "meses": True})
        self.assertEqual([len(e[2]) for e in conn.log if e[0] == "executemany"], [3])

    def test_una_liquidacion_con_resultado_si_se_escribe(self):
        mes = _mes("liquidacion", 2)
        conn = _Conexion(RESPUESTAS_GUARDAR)
        self.assertEqual(_guardar(conn, [mes]), {"liquidacion": True, "meses": True})
        filas = next(e for e in conn.log if e[0] == "executemany")[2]
        self.assertTrue(all(f[8] == RES_INCLUIDA for f in filas))   # columna resultado


class ReglasDelCruceDejanDestinoParaTodas(DestinosPorRegla):
    """Reusa el armado de casos de DestinosPorRegla: toda guia, sea cual sea la regla que la toque, queda con destino."""

    def test_toda_guia_tiene_destino_y_arma_filas_de_liquidacion(self):
        from tests.test_jetsmart_duplicados import _export_con
        d = self._run(export_guias=[10, 11, 12, 13], ariel_guias=[10, 12, 20, 30], prev_ariel=[12], prev_export=[20])
        self.assertEqual(set(d), {10, 11, 12, 13, 20, 30})
        self.assertTrue(all(v[0] in RESULTADOS_VALIDOS for v in d.values()))


def _hay(*archivos):
    return unittest.skipUnless(all((FIX / a).exists() for a in archivos), "faltan archivos reales en tests/fixtures/jetsmart/")


class MesesReales(unittest.TestCase):
    @_hay("Guias_Aerea_10-2026.xlsx", "Guias_Aerea_09-2026.xlsx", "Ariel_Rendicion_octubre_2026.xlsx", "Ariel_Rendicion_septiembre_2026.xlsx")
    def test_octubre_2026(self):
        g_oct, g_sep = cargar_guias_sistema(FIX / "Guias_Aerea_10-2026.xlsx"), cargar_guias_sistema(FIX / "Guias_Aerea_09-2026.xlsx")
        a_oct, a_sep = leer_ariel(FIX / "Ariel_Rendicion_octubre_2026.xlsx", cruda=True), leer_ariel(FIX / "Ariel_Rendicion_septiembre_2026.xlsx", cruda=True)
        r = cruzar_jetsmart(g_oct, a_oct.guias, ("OCT", "26"), g_sep, a_sep.guias, ariel_repetidas_identicas=a_oct.repetidas_identicas)
        filas = _verificar(self, g_oct, a_oct.guias, r.destinos)
        # incluidas y pendientes de declarar, todas con resultado
        self.assertEqual({f["resultado"] for f in filas}, {RES_INCLUIDA, RES_PENDIENTE})
        self.assertEqual(len(filas), 1944 + 1953)

    @_hay("PRUEBA_LIQUIDACION_08-2026.xlsx", "Ariel_Rendicion_agosto_2026.xlsx", "LIQUIDACION_07-2026.xlsx")
    def test_agosto_2026_con_duplicada_diferencias_y_revision_manual(self):
        # Agosto trae de todo: 97479 repetida con datos distintos, dos diferencias de centavos,
        # guias sin confirmar y 21 ya declaradas el mes anterior.
        r = cruzar_jetsmart(
            load_jetsmart_export(FIX / "PRUEBA_LIQUIDACION_08-2026.xlsx"),
            leer_ariel(FIX / "Ariel_Rendicion_agosto_2026.xlsx", cruda=True).guias, ("AGO", "26"),
            load_jetsmart_export(FIX / "LIQUIDACION_07-2026.xlsx"), load_ariel(FIX / "LIQUIDACION_07-2026.xlsx")[0],
        )
        filas = _verificar(
            self, load_jetsmart_export(FIX / "PRUEBA_LIQUIDACION_08-2026.xlsx"),
            leer_ariel(FIX / "Ariel_Rendicion_agosto_2026.xlsx", cruda=True).guias, r.destinos,
        )
        self.assertEqual(r.destinos[97479][0], RES_EXCLUIDA)            # repetida con datos distintos
        self.assertEqual(r.destinos[96889][0], RES_INCLUIDA)            # diferencia de centavos: entra con valores de Ariel
        self.assertEqual({f["resultado"] for f in filas} - RESULTADOS_VALIDOS, set())
        self.assertIn(RES_EXCLUIDA, {f["resultado"] for f in filas})

    @_hay("PRUEBA_LIQUIDACION_09-2026.xlsx", "PRUEBA_LIQUIDACION_08-2026.xlsx", "Ariel_Rendicion_septiembre_2026.xlsx", "Ariel_Rendicion_agosto_2026.xlsx")
    def test_septiembre_2026_con_guias_declaradas_dos_veces(self):
        export_sep = load_jetsmart_export(FIX / "PRUEBA_LIQUIDACION_09-2026.xlsx")
        ariel_sep = leer_ariel(FIX / "Ariel_Rendicion_septiembre_2026.xlsx", cruda=True).guias
        r = cruzar_jetsmart(
            export_sep, ariel_sep, ("SEP", "26"),
            load_jetsmart_export(FIX / "PRUEBA_LIQUIDACION_08-2026.xlsx"),
            leer_ariel(FIX / "Ariel_Rendicion_agosto_2026.xlsx", cruda=True).guias,
        )
        _verificar(self, export_sep, ariel_sep, r.destinos)
        for g in (98172, 98177, 98182, 98188):                          # declaradas en agosto y otra vez en septiembre
            self.assertEqual(r.destinos[g][0], RES_EXCLUIDA)


if __name__ == "__main__":
    unittest.main()
