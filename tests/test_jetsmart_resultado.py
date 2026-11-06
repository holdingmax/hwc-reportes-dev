"""Pantalla de RESULTADO de JetSmart: estado general, numeros clave, boton y detalle.

Correr con:
    python -m unittest tests.test_jetsmart_resultado -v

Usa los archivos reales de octubre y septiembre 2026 en
tests/fixtures/jetsmart/ (no versionado): si faltan, se saltea. Nunca
escribe en la base: se quita DATABASE_URL y el guardado y la lectura de la
version se reemplazan por dobles.
"""

import dataclasses
import os
import re
import unittest
from datetime import datetime, timezone
from unittest import mock

from streamlit.testing.v1.element_tree import Block

from src.jetsmart_cruce import cruzar_jetsmart
from src.jetsmart_mes_anterior import LecturaMesAnterior, ResumenMesAnterior
from tests.test_jetsmart_pantalla import ARCHIVOS, _app

CLAVES = ["iva_oct", "ariel_sep", "guias_oct", "guias_sep", "ariel_oct"]
GUARDADO_OK = {"liquidacion": True, "meses": True}


def _sin_excepciones(*args, **kwargs):
    """El cruce real, pero sin nada para revisar (para ver el caso verde)."""
    resultado = cruzar_jetsmart(*args, **kwargs)
    resumen = {**resultado.resumen, "excepciones": 0, "excepciones_por_tipo": {}, "ingreso_fuera_por_revision": 0.0}
    return dataclasses.replace(resultado, excepciones=resultado.excepciones.iloc[0:0], resumen=resumen)


def _lectura_version(version: int):
    def leer(period, con_guias=False):
        if period == ("OCT", "26"):
            resumen = ResumenMesAnterior(period, version, datetime(2026, 11, 4, 15, 0, tzinfo=timezone.utc), "liquidacion")
            return LecturaMesAnterior("ok", resumen)
        return LecturaMesAnterior("sin_base")
    return leer


def _es_detalle(bloque) -> bool:
    """Los expanders del detalle (en este AppTest salen como bloques con rotulo).

    "Cambiar a mano" es del paso 3 (datos del periodo), no del detalle del resultado."""
    rotulo = getattr(bloque, "label", None)
    return rotulo is not None and rotulo != "Cambiar a mano" and hasattr(getattr(bloque, "proto", None), "expanded")


def _detalles(at) -> list:
    """Los expanders del DETALLE de la liquidacion, en orden (sin el del Historial)."""
    encontrados = []

    def recorrer(bloque):
        for hijo in bloque.children.values():
            if isinstance(hijo, Block):
                if _es_detalle(hijo) and not hijo.label.startswith("Ver historial"):
                    encontrados.append(hijo)
                recorrer(hijo)

    recorrer(at.main)
    return encontrados


def _htmls(at, clase: str) -> list[str]:
    return [m.value for m in at.markdown if clase in m.value and "<style>" not in m.value]


@unittest.skipUnless(all(p.exists() for p in ARCHIVOS.values()), "faltan los archivos reales")
class ResultadoJetSmart(unittest.TestCase):
    def setUp(self):
        self._env = mock.patch.dict(os.environ)
        self._env.start()
        os.environ.pop("DATABASE_URL", None)
        self.addCleanup(self._env.stop)

    def _procesar(
        self, guardado=GUARDADO_OK, version: int | None = 3, sin_excepciones: bool = False,
        vuelos: tuple[int, int] | None = None,
    ):
        patches = [mock.patch("src.db.guardar_liquidacion_jetsmart", return_value=guardado)]
        if version is not None:
            patches.append(mock.patch("src.db.obtener_mes_anterior_jetsmart", side_effect=_lectura_version(version)))
        if sin_excepciones:
            patches.append(mock.patch("src.jetsmart_cruce.cruzar_jetsmart", side_effect=_sin_excepciones))
        for p in patches:
            p.start()
            self.addCleanup(p.stop)
        at = _app(CLAVES)
        at.number_input[0].set_value(1485.0)
        if vuelos is not None:  # si no, quedan los que precarga el Excel de Ariel (EZE 49, MDZ 27)
            at.number_input(key="js_vuelos_EZE").set_value(vuelos[0])
            at.number_input(key="js_vuelos_MDZ").set_value(vuelos[1])
        at.run()
        at.button[0].click()
        at.run()
        self.assertEqual(len(at.exception), 0, [e.value for e in at.exception])
        return at

    def test_banner_amarillo_si_hay_guias_para_revisar(self):
        at = self._procesar()
        banner = _htmls(at, "hwc-banner-warn")
        self.assertEqual(len(banner), 1)
        self.assertIn("Liquidación lista, pero hay 10 guías para revisar", banner[0])
        self.assertEqual(_htmls(at, "hwc-banner-ok"), [])

    def test_banner_verde_si_no_hay_nada_para_revisar(self):
        at = self._procesar(sin_excepciones=True)
        banner = _htmls(at, "hwc-banner-ok")
        self.assertEqual(len(banner), 1)
        self.assertIn("Liquidación de octubre 2026 lista", banner[0])
        self.assertEqual(_htmls(at, "hwc-banner-warn"), [])

    def test_debajo_del_banner_la_version_y_la_fecha(self):
        at = self._procesar(version=3)
        self.assertIn("Versión 3, guardada el 4/11", _htmls(at, "hwc-banner-sub")[0])

    def test_si_no_se_guardo_lo_dice_y_pide_descargar(self):
        at = self._procesar(guardado={"liquidacion": False, "meses": False}, version=None)
        sub = _htmls(at, "hwc-banner-sub")[0]
        self.assertIn("No se guardó en el Historial", sub)
        self.assertIn("hwc-banner-sub-warn", sub)
        self.assertEqual(len(at.get("download_button")), 1)

    def test_tres_tarjetas_con_cifra_grande_y_etiqueta(self):
        at = self._procesar()
        tarjetas = _htmls(at, "hwc-numcard")
        self.assertEqual(len(tarjetas), 3)
        total, guias, revisar = tarjetas
        self.assertIn("hwc-numcard-hero", total)
        self.assertIn("Total a entregar a WCS", total)
        self.assertIn("$ 66.513.964", total)
        self.assertIn('<div class="hwc-numcard-value">1.953</div>', guias)
        self.assertIn("Guías liquidadas", guias)
        self.assertIn("hwc-numcard-warn", revisar)
        self.assertIn('<div class="hwc-numcard-value">10</div>', revisar)
        self.assertIn("Para revisar", revisar)

    def test_tarjeta_para_revisar_verde_si_es_cero(self):
        at = self._procesar(sin_excepciones=True)
        revisar = _htmls(at, "hwc-numcard")[2]
        self.assertIn("hwc-numcard-ok", revisar)
        self.assertIn('<div class="hwc-numcard-value">0</div>', revisar)

    def test_cvlp_va_como_linea_secundaria_debajo_de_las_tarjetas(self):
        at = self._procesar()
        secundaria = _htmls(at, "hwc-secondary")
        self.assertEqual(len(secundaria), 1)
        self.assertIn("CVLP", secundaria[0])
        self.assertIn("$ 81.044.006,16", secundaria[0])
        self.assertEqual(len(_htmls(at, "hwc-numcard")), 3)  # no es una cuarta tarjeta

    def test_un_solo_boton_descargar_excel_primario(self):
        at = self._procesar()
        botones = at.get("download_button")
        self.assertEqual(len(botones), 1)
        self.assertEqual(botones[0].proto.label, "Descargar Excel")
        self.assertEqual(botones[0].proto.type, "primary")

    def test_orden_banner_tarjetas_boton(self):
        at = self._procesar()
        orden = []

        def recorrer(bloque):
            for hijo in bloque.children.values():
                if isinstance(hijo, Block):
                    if _es_detalle(hijo) and not hijo.label.startswith("Ver historial"):
                        orden.append("detalle")
                    recorrer(hijo)
                    continue
                valor = getattr(hijo, "value", "")
                valor = valor if isinstance(valor, str) and "<style>" not in valor else ""
                if hijo.type == "download_button":
                    orden.append("boton")
                elif "hwc-banner-warn" in valor or "hwc-banner-ok" in valor:
                    orden.append("banner")
                elif "hwc-numcard-hero" in valor:
                    orden.append("tarjetas")
                elif 'class="hwc-review"' in valor:
                    orden.append("para_revisar")

        recorrer(at.main)
        self.assertEqual(orden, ["banner", "tarjetas", "para_revisar", "boton", "detalle", "detalle", "detalle"])

    def test_para_revisar_lista_cada_guia_con_que_hacer(self):
        at = self._procesar()
        bloque = _htmls(at, 'class="hwc-review"')
        self.assertEqual(len(bloque), 1)
        bloque = bloque[0]
        self.assertIn("Para revisar: 10 guías", bloque)
        for columna in ("Guía", "Qué pasa", "¿Entra?", "Qué hacer"):
            self.assertIn(f"<th>{columna}</th>", bloque)
        self.assertEqual(bloque.count("<tr><td"), 10)  # una fila por guia
        self.assertEqual(bloque.count("Esperá a que Ariel la declare"), bloque.count("Ariel todavía no la declaró"))
        # Los numeros de guia, enteros y sin decimales.
        numeros = re.findall(r'<td class="hwc-num">([^<]*)</td>', bloque)
        self.assertEqual(len(numeros), 10)
        self.assertTrue(all(n.isdigit() for n in numeros), numeros)
        # Palabras de Anita: el cruce se explica, y no aparece la jerga "excepcion".
        self.assertIn("el cruce de los dos listados", bloque)
        self.assertNotIn("xcepci", bloque)

    def test_sin_guias_para_revisar_no_hay_bloque(self):
        at = self._procesar(sin_excepciones=True)
        self.assertEqual(_htmls(at, 'class="hwc-review"'), [])

    def test_la_tabla_de_excepciones_ya_no_se_repite_mas_abajo(self):
        at = self._procesar()
        self.assertFalse(any("Excepciones del cruce" in m.value for m in at.markdown))
        # La tabla completa solo vive dentro del expander cerrado del cruce, no suelta en la pantalla.
        cruce = _detalles(at)[0]
        self.assertEqual(len(cruce.dataframe), 1)
        self.assertEqual(
            [d for d in at.dataframe if "Tipo" in d.value.columns and d.key != cruce.dataframe[0].key], [],
        )

    def test_el_detalle_son_tres_expanders_cerrados(self):
        at = self._procesar()
        detalle = _detalles(at)
        self.assertEqual(
            [d.label for d in detalle],
            ["Cómo se armó el cruce", "Liquidación: comisiones, GHA y IVA", "Tabla de guías (1.953)"],
        )
        self.assertFalse(any(d.proto.expanded for d in detalle))

    def test_el_detalle_tiene_el_cruce_la_liquidacion_y_la_tabla(self):
        at = self._procesar()
        cruce, liquidacion, tabla = _detalles(at)
        cruce_txt = " ".join(m.value for m in cruce.markdown) + " ".join(c.value for c in cruce.caption)
        self.assertIn("El cruce compara las guías del sistema con lo que declaró Ariel", cruce_txt)
        self.assertIn("Mes anterior: septiembre 2026", cruce_txt)
        for fila in ("Guías que entran a la liquidación", "Coinciden exacto", "Ya declaradas el mes anterior",
                     "Anuladas por Ariel (no se cuentan)", "Filas repetidas de Ariel"):
            self.assertIn(fila, cruce_txt)
        self.assertNotIn("Matchean", cruce_txt)  # jerga
        self.assertIn("Total a entregar a WCS", " ".join(m.value for m in liquidacion.markdown))
        self.assertEqual(len(tabla.dataframe), 1)

    def test_tabla_de_guias_con_numero_entero(self):
        at = self._procesar()
        (df,) = [d.value for d in at.dataframe if "Nro guía" in d.value.columns and "TOTAL" in d.value.columns]
        self.assertEqual(len(df), 1953)
        self.assertTrue(all(v.isdigit() for v in df["Nro guía"]), "la guia debe verse sin decimales")

    def test_datos_usados_en_una_linea_chica(self):
        at = self._procesar(vuelos=(49, 27))
        lineas = [c.value for c in at.caption if c.value.startswith("Datos usados:")]
        self.assertEqual(len(lineas), 1)
        self.assertIn("tipo de cambio $ 1.485,00", lineas[0])
        self.assertIn("EZE 49 vuelos × 17,5 USD = 857,5 USD", lineas[0])
        self.assertIn("MDZ 27 vuelos × 35 USD = 945 USD", lineas[0])

    def test_los_avisos_importantes_siguen_a_la_vista(self):
        at = self._procesar()
        avisos = " ".join(w.value for w in at.warning)
        self.assertIn("20 filas repetidas idénticas", avisos)
        # ...y no quedaron escondidos dentro de un expander cerrado.
        self.assertFalse(any("repetidas idénticas" in w.value for d in _detalles(at) for w in d.warning))


if __name__ == "__main__":
    unittest.main()
