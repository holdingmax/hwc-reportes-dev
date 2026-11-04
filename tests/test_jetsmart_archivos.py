"""Reconocimiento automatico de los archivos de JetSmart (src/jetsmart_archivos.py).

Correr con:
    python -m unittest tests.test_jetsmart_archivos -v

Usa los archivos reales de octubre 2026 en tests/fixtures/jetsmart/ (no
versionado): si faltan, esos tests se saltean. Los de errores usan archivos
chicos armados en el momento.
"""

import io
import itertools
import unittest
from pathlib import Path

import pandas as pd
from openpyxl import Workbook

from src.config import COL_JS_GUIA, COL_JS_KGS, COL_JS_PRIORIDAD
from src.jetsmart_archivos import (
    TIPO_ARIEL,
    TIPO_DESCONOCIDO,
    TIPO_GUIAS,
    TIPO_IVA,
    TIPO_LISTADO,
    TIPO_TRABAJO,
    armar_conjunto,
    cargar_ariel,
    cargar_guias_sistema,
    detectar_archivo,
)
from src.jetsmart_builder import load_jetsmart_export

FIX = Path(__file__).resolve().parent / "fixtures" / "jetsmart"
GUIAS = FIX / "Guias_Aerea_10-2026.xlsx"
GUIAS_SEP = FIX / "Guias_Aerea_09-2026.xlsx"
LISTADO = FIX / "Listado_Facturacion_JM_10-2026.xlsx"
ARIEL_OCT = FIX / "Ariel_Rendicion_octubre_2026.xlsx"
ARIEL_SEP = FIX / "Ariel_Rendicion_septiembre_2026.xlsx"
ARIEL_AGO = FIX / "Ariel_Rendicion_agosto_2026.xlsx"
IVA = FIX / "IVA_VENTAS_HW_10-2026.xlsx"
TRABAJO_JUL = FIX / "LIQUIDACION_07-2026.xlsx"
LIQ_ARMADA = FIX / "LIQ_ECS_07-2026_00000005.xlsx"
OCT = ("OCT", "26")
SEP = ("SEP", "26")

hay_octubre = unittest.skipUnless(all(p.exists() for p in (GUIAS, LISTADO, ARIEL_OCT, IVA)), "faltan los archivos reales de octubre")


def _detectar(path: Path):
    with open(path, "rb") as f:
        return detectar_archivo(f, path.name)


def _xlsx_chico(hojas: dict[str, list[list]]) -> io.BytesIO:
    wb = Workbook()
    wb.remove(wb.active)
    for nombre, filas in hojas.items():
        ws = wb.create_sheet(nombre)
        for fila in filas:
            ws.append(fila)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


@hay_octubre
class DeteccionArchivosReales(unittest.TestCase):
    def test_guias_del_sistema(self):
        d = _detectar(GUIAS)
        self.assertEqual((d.tipo, d.cantidad, d.periodo), (TIPO_GUIAS, 1944, OCT))
        self.assertIn("Detecté: Guías del sistema de octubre 2026, 1.944 guías", d.resumen)

    @unittest.skipUnless(GUIAS_SEP.exists(), "faltan las guías de septiembre")
    def test_guias_de_septiembre(self):
        d = _detectar(GUIAS_SEP)
        self.assertEqual((d.tipo, d.cantidad, d.periodo), (TIPO_GUIAS, 1984, SEP))

    def test_listado_de_facturacion(self):
        d = _detectar(LISTADO)
        self.assertEqual((d.tipo, d.cantidad, d.periodo), (TIPO_LISTADO, 1905, OCT))

    def test_rendicion_de_ariel_cuenta_guias_unicas_y_avisa_anuladas(self):
        d = _detectar(ARIEL_OCT)
        # 1.973 filas validas, 20 repetidas identicas -> 1.953 guias distintas; 8 anuladas.
        self.assertEqual((d.tipo, d.cantidad, d.periodo), (TIPO_ARIEL, 1953, OCT))
        self.assertIn("8 anuladas", d.resumen)
        self.assertIn("20 filas repetidas idénticas, se cuentan una sola vez", d.resumen)

    def test_libro_iva(self):
        d = _detectar(IVA)
        self.assertEqual((d.tipo, d.periodo), (TIPO_IVA, OCT))
        self.assertGreater(d.cantidad, 4000)
        self.assertIn("comprobantes", d.resumen)

    @unittest.skipUnless(TRABAJO_JUL.exists(), "falta el archivo de trabajo de julio")
    def test_archivo_de_trabajo_se_reconoce_como_tal(self):
        d = _detectar(TRABAJO_JUL)
        self.assertEqual((d.tipo, d.periodo), (TIPO_TRABAJO, ("JUL", "26")))

    @unittest.skipUnless(LIQ_ARMADA.exists(), "falta la liquidacion armada de julio")
    def test_liquidacion_ya_armada_se_explica(self):
        d = _detectar(LIQ_ARMADA)
        self.assertEqual(d.tipo, TIPO_DESCONOCIDO)
        self.assertIn("liquidación ya armada", d.problema)

    def test_el_orden_y_el_nombre_del_archivo_no_importan(self):
        # Se reconoce por columnas: renombrar el archivo no cambia nada.
        with open(GUIAS, "rb") as f:
            d = detectar_archivo(f, "cualquier_cosa.xlsx")
        self.assertEqual(d.tipo, TIPO_GUIAS)


@hay_octubre
class GuiasDelSistemaComoExport(unittest.TestCase):
    """Las guias del sistema, mapeadas al formato interno, tienen que dar lo
    mismo que el listado en las guias que comparten, y cruzar contra Ariel."""

    def test_mismas_columnas_que_el_listado(self):
        self.assertEqual(list(cargar_guias_sistema(GUIAS).columns), list(load_jetsmart_export(LISTADO).columns))

    def test_guias_y_listado_coinciden_en_kg_y_prioridad(self):
        g = cargar_guias_sistema(GUIAS).set_index(COL_JS_GUIA)
        l = load_jetsmart_export(LISTADO).set_index(COL_JS_GUIA)
        comunes = g.index.intersection(l.index)
        self.assertEqual(len(comunes), 1857)
        self.assertEqual((g.loc[comunes, COL_JS_KGS].astype(float) != l.loc[comunes, COL_JS_KGS].astype(float)).sum(), 0)
        self.assertEqual((g.loc[comunes, COL_JS_PRIORIDAD].astype(float) != l.loc[comunes, COL_JS_PRIORIDAD].astype(float)).sum(), 0)

    def test_solo_se_leen_las_columnas_que_usa_anita(self):
        # Cliente y destino salen de Ariel, no del sistema: quedan vacios.
        g = cargar_guias_sistema(GUIAS)
        self.assertTrue(g["Cliente"].isna().all() and g["Aero. Dstino"].isna().all())
        self.assertEqual(set(g["Empresa"]), {"Jetsmart"})

    def test_alcanza_con_las_seis_columnas_de_anita(self):
        # Un export con SOLO #, Kgs., $ Prioridad, Aero.Ori., Estado y Cre. se reconoce y se lee.
        chico = _xlsx_chico({"Sheet1": [
            ["#", "Kgs.", "$ Prioridad", "Aero.Ori.", "Estado", "Cre."],
            [102178, 0.5, 12000.0, "AEP", "Entregada", "31/10/2026"],
            [102177, 11.0, 33055.0, "NQN", "En vuelo", "31/10/2026"],
        ]})
        d = detectar_archivo(chico, "chico.xlsx")
        self.assertEqual((d.tipo, d.cantidad, d.periodo), (TIPO_GUIAS, 2, OCT))
        self.assertEqual(list(cargar_guias_sistema(chico)[COL_JS_GUIA]), [102178, 102177])

    def test_la_clave_de_cruce_es_la_columna_numeral_no_nro(self):
        # "Nro." trae "CRD-102178"; la guia de Ariel es 102178.
        g = cargar_guias_sistema(GUIAS)
        ariel = cargar_ariel(ARIEL_OCT, TIPO_ARIEL).guias
        comunes = set(pd.to_numeric(g[COL_JS_GUIA])) & set(ariel["Nro de guía"])
        self.assertEqual(len(comunes), 1934)


class ErroresEnLenguajeDeNegocio(unittest.TestCase):
    def _problema(self, buf) -> str:
        d = detectar_archivo(buf, "raro.xlsx")
        self.assertEqual(d.tipo, TIPO_DESCONOCIDO)
        for tecnico in ("ValueError", "Traceback", "KeyError", "BD", "ARIEL", "LIQUIDACION ECS"):
            self.assertNotIn(tecnico, d.problema)
        return d.problema

    def test_excel_cualquiera_lista_lo_que_se_acepta(self):
        msg = self._problema(_xlsx_chico({"Hoja1": [["a", "b"], [1, 2]]}))
        self.assertIn("rendición de Ariel", msg)
        self.assertIn("Hoja1", msg)

    def test_archivo_que_no_es_excel(self):
        msg = self._problema(io.BytesIO(b"esto no es un excel"))
        self.assertIn("No pude abrir", msg)

    def test_excel_vacio(self):
        self._problema(_xlsx_chico({"Hoja1": []}))


def _ar(nombre, tipo, periodo, cantidad=10):
    from src.jetsmart_archivos import ArchivoDetectado
    return ArchivoDetectado(nombre, tipo, cantidad=cantidad, periodo=periodo)


class ConjuntoDeArchivos(unittest.TestCase):
    def test_sin_archivos_pide_lo_basico(self):
        c = armar_conjunto([])
        self.assertFalse(c.listo)
        self.assertIn("guías del mes", c.faltantes[0])

    def test_falta_el_mes_anterior_y_lo_dice(self):
        c = armar_conjunto([_ar("g.xlsx", TIPO_GUIAS, OCT), _ar("a.xlsx", TIPO_ARIEL, OCT)])
        self.assertEqual(c.periodo, OCT)
        self.assertFalse(c.listo)
        texto = " ".join(c.faltantes)
        self.assertIn("septiembre 2026", texto)
        self.assertNotIn("octubre", texto)

    def test_conjunto_completo_en_cualquier_orden(self):
        archivos = [
            _ar("g_oct", TIPO_GUIAS, OCT), _ar("a_oct", TIPO_ARIEL, OCT),
            _ar("g_sep", TIPO_GUIAS, SEP), _ar("a_sep", TIPO_ARIEL, SEP), _ar("iva", TIPO_IVA, OCT),
        ]
        for orden in itertools.permutations(archivos):
            c = armar_conjunto(list(orden))
            self.assertTrue(c.listo, orden)
            self.assertEqual(
                (c.export_mes.nombre, c.ariel_mes.nombre, c.export_anterior.nombre, c.ariel_anterior.nombre),
                ("g_oct", "a_oct", "g_sep", "a_sep"),
            )

    def test_guias_del_sistema_ganan_al_listado(self):
        c = armar_conjunto([_ar("listado", TIPO_LISTADO, OCT), _ar("g", TIPO_GUIAS, OCT), _ar("a", TIPO_ARIEL, OCT)])
        self.assertEqual(c.export_mes.nombre, "g")
        self.assertFalse(any("listado de facturación como guías" in a for a in c.avisos))

    def test_solo_listado_se_usa_pero_con_aviso(self):
        c = armar_conjunto([_ar("listado", TIPO_LISTADO, OCT), _ar("a", TIPO_ARIEL, OCT)])
        self.assertEqual(c.export_mes.nombre, "listado")
        self.assertTrue(any("solo trae las guías ya facturadas" in a.lower() or "solo trae las guías ya facturadas" in a for a in c.avisos))

    def test_archivo_de_trabajo_cubre_guias_y_ariel_del_mes_anterior(self):
        c = armar_conjunto([_ar("g", TIPO_GUIAS, OCT), _ar("a", TIPO_ARIEL, OCT), _ar("trabajo_sep", TIPO_TRABAJO, SEP)])
        self.assertTrue(c.listo)
        self.assertEqual((c.export_anterior.nombre, c.ariel_anterior.nombre), ("trabajo_sep", "trabajo_sep"))

    def test_dos_archivos_del_mismo_tipo_se_avisa_cual_se_usa(self):
        c = armar_conjunto([_ar("g1", TIPO_GUIAS, OCT), _ar("g2", TIPO_GUIAS, OCT), _ar("a", TIPO_ARIEL, OCT)])
        self.assertEqual(c.export_mes.nombre, "g2")
        self.assertTrue(any("g2" in a and "g1" in a for a in c.avisos))

    def test_archivo_de_otro_mes_no_se_ignora_en_silencio(self):
        c = armar_conjunto([_ar("a_oct", TIPO_ARIEL, OCT), _ar("a_ago", TIPO_ARIEL, ("AGO", "26"))])
        self.assertTrue(any("a_ago" in a and "agosto 2026" in a for a in c.avisos))

    def test_archivo_no_reconocido_va_a_errores_sin_bloquear(self):
        from src.jetsmart_archivos import ArchivoDetectado
        raro = ArchivoDetectado("raro.xlsx", TIPO_DESCONOCIDO, problema="No reconozco este archivo.")
        archivos = [raro, _ar("g_oct", TIPO_GUIAS, OCT), _ar("a_oct", TIPO_ARIEL, OCT), _ar("g_sep", TIPO_GUIAS, SEP), _ar("a_sep", TIPO_ARIEL, SEP)]
        c = armar_conjunto(archivos)
        self.assertTrue(c.listo)
        self.assertEqual(len(c.errores), 1)
        self.assertIn("raro.xlsx", c.errores[0])

    @hay_octubre
    def test_octubre_real_con_archivos_reales(self):
        archivos = [_detectar(p) for p in (IVA, ARIEL_OCT, LISTADO, GUIAS)]
        c = armar_conjunto(archivos)
        self.assertEqual(c.periodo, OCT)
        self.assertEqual(c.export_mes.tipo, TIPO_GUIAS)   # no el listado
        self.assertEqual(c.ariel_mes.tipo, TIPO_ARIEL)
        self.assertEqual(c.listado.tipo, TIPO_LISTADO)
        self.assertEqual(c.libro_iva.tipo, TIPO_IVA)
        self.assertEqual(len(c.faltantes), 2)             # solo el mes anterior


if __name__ == "__main__":
    unittest.main()
