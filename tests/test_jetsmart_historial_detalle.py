"""Detalle de una liquidacion de JetSmart en el Historial.

Correr con:
    python -m unittest tests.test_jetsmart_historial_detalle -v

Sin archivos reales ni base: simula lo que devuelve la base y lo muestra con
la pantalla real (AppTest). Cubre:
  1. numeros de guia enteros (tablas de la pantalla y Excel descargable);
  2. "—" (y no "None") cuando una guia no tiene dato de Ariel o del sistema;
  3. el Historial usa los mismos componentes que el resultado nuevo.
"""

import io
import os
import re
import unittest
from datetime import datetime
from pathlib import Path
from unittest import mock

import pandas as pd
from openpyxl import load_workbook
from streamlit.testing.v1 import AppTest

from src import jetsmart_builder as b
from src.jetsmart_builder import build_jetsmart_resumen_guardado, guias_guardadas, write_jetsmart_liquidacion
from src.jetsmart_cruce import EXC_DIFERENCIA, EXC_PENDIENTE

APP = Path(__file__).resolve().parent.parent / "app.py"

# Como vuelven de la base: con ruido decimal.
RUIDO = [100186.00000000000001, 102141.00000000000001, 102176.99999999999]
GUIAS = [
    {"Nro guía": g, "Kg": 10.0, "GRAVADO": 1000.0, "IVA": 210.0, "NO GRAV": 0.0, "TOTAL": 1210.0, "Origen": "AEP"}
    for g in RUIDO
]
EXCEPCIONES = [
    {  # pendiente: no esta en Ariel -> sin dato de Ariel
        "Tipo": EXC_PENDIENTE, "Nro guía": RUIDO[0], "Fecha Ariel": None, "Origen": "AEP",
        "Kg export": 12.5, "Kg Ariel": None, "Ingreso export": 3000.0, "Ingreso Ariel": None,
        "¿Entra a la liquidación?": "No", "Detalle": "Esta en el sistema pero Ariel no la declaro.",
    },
    {
        "Tipo": EXC_DIFERENCIA, "Nro guía": RUIDO[2], "Fecha Ariel": "2026-10-03", "Origen": "COR",
        "Kg export": 5.0, "Kg Ariel": 6.0, "Ingreso export": 100.0, "Ingreso Ariel": 120.0,
        "¿Entra a la liquidación?": "Sí, con valores de Ariel", "Detalle": "Difiere Kg.",
    },
]
ULTIMO_DIA = [{"Nro guía": RUIDO[1], "Fecha Ariel": "2026-10-31", "Origen": "AEP", "Kg Ariel": 3.0, "Ingreso Ariel": 50.0}]
RESUMEN_CRUCE = {
    "excepciones": 2, "incluidas": 3, "motivos_inclusion": {"match": 1}, "ya_declaradas_mes_anterior": 0,
    "ingreso_fuera_por_revision": 0, "ultimo_dia": "31/10", "ariel_repetidas_identicas": [],
}


class _Historial(unittest.TestCase):
    def _abrir(self, cruce="con_cruce", parametros=None):
        fila = pd.DataFrame([{
            "id": 7, "carga_id": 7, "aerolinea": "jetsmart", "estacion": "AR", "tipo_cargo": "jetsmart_liquidacion",
            "periodo_mes": "OCT", "periodo_anio": "26", "periodo": datetime(2026, 10, 1), "cantidad_filas": 3,
            "monto_total": 1000.0, "es_vigente": True, "generado_en": pd.Timestamp("2026-11-04 12:00"),
        }])
        detalle = {
            "guias": GUIAS, "parametros": parametros or {"tipo_cambio": 1000.0, "vuelos_inter": {}},
            "cruce": None if cruce is None else {
                "resumen": RESUMEN_CRUCE, "excepciones": EXCEPCIONES, "incluidas_ultimo_dia": ULTIMO_DIA,
            },
        }
        env = mock.patch.dict(os.environ)
        env.start()
        os.environ.pop("DATABASE_URL", None)
        self.addCleanup(env.stop)
        for nombre, valor in (
            ("obtener_filtros_historial", {"aerolineas": ["jetsmart"], "periodos": [(datetime(2026, 10, 1), "OCT", "26")], "estaciones": ["AR"]}),
            ("obtener_historial", fila),
            ("obtener_detalle_totales", detalle),
        ):
            patch = mock.patch(f"src.db.{nombre}", return_value=valor)
            patch.start()
            self.addCleanup(patch.stop)
        at = AppTest.from_file(str(APP), default_timeout=120)
        at.run()
        self.assertEqual(len(at.exception), 0, [e.value for e in at.exception])
        return at


def _htmls(at):
    return [m.value for m in at.markdown if isinstance(m.value, str)]


class NumerosDeGuiaEnteros(_Historial):
    """Punto 1: ninguna tabla muestra decimales en el numero de guia."""

    def test_las_tablas_del_historial_muestran_la_guia_entera(self):
        at = self._abrir()
        con_guia = [df.value for df in at.dataframe if "Nro guía" in df.value.columns]
        self.assertGreaterEqual(len(con_guia), 3, "faltan tablas: guias, excepciones y ultimo dia")
        for tabla in con_guia:
            for v in tabla["Nro guía"]:
                self.assertRegex(str(v), r"^\d+$", f"guia con decimales: {v!r}")
        todas = {str(v) for t in con_guia for v in t["Nro guía"]}
        self.assertEqual(todas, {"100186", "102141", "102177"})

    def test_la_guia_se_redondea_no_se_corta(self):
        at = self._abrir()
        guias = next(df.value for df in at.dataframe if "TOTAL" in df.value.columns)
        self.assertIn("102177", list(guias["Nro guía"]))  # 102176.99999999999, no 102176

    def test_el_excel_trae_las_guias_enteras(self):
        # Armado real: el Excel recibe el cruce con guias de ruido y debe escribirlas enteras.
        guias = guias_guardadas(GUIAS)
        resumen = build_jetsmart_resumen_guardado(guias, {"tipo_cambio": 1000.0, "vuelos_inter": {}})
        buffer = io.BytesIO()
        cruce = {
            "excepciones": pd.DataFrame(EXCEPCIONES), "incluidas_ultimo_dia": pd.DataFrame(ULTIMO_DIA),
            "resumen": {**RESUMEN_CRUCE, "ariel_repetidas_identicas": [RUIDO[0]]},
        }
        write_jetsmart_liquidacion(guias, resumen, ("OCT", "26"), buffer, cruce=cruce)
        wb = load_workbook(io.BytesIO(buffer.getvalue()))
        for hoja in ("GUIAS", "EXCEPCIONES CRUCE", "INCLUIDAS ULTIMO DIA", "REPETIDAS EN ARIEL"):
            col = [c.value for c in wb[hoja][1]].index("Nro guía") + 1
            valores = [wb[hoja].cell(row=r, column=col).value for r in range(2, wb[hoja].max_row + 1)]
            self.assertTrue(valores, hoja)
            valores = [v for v in valores if v != "TOTAL"]  # la hoja GUIAS cierra con una fila de totales
            self.assertTrue(all(isinstance(v, int) for v in valores), f"{hoja}: {valores}")

    def test_la_liquidacion_recien_hecha_arma_guias_enteras(self):
        df = pd.DataFrame({
            "Nro guía": RUIDO, "Fecha creación": "2026-10-01", "Cliente": "X", "Origen": "AEP", "Destino": "BRC",
            "Estado": "OK", "Kgs": 1.0, "$ Prioridad": 100.0,
        })
        df = df.rename(columns={
            "Nro guía": b.COL_JS_GUIA, "Fecha creación": b.COL_JS_CREACION, "Cliente": b.COL_JS_CLIENTE,
            "Origen": b.COL_JS_ORIGEN, "Destino": b.COL_JS_DESTINO, "Estado": b.COL_JS_ESTADO,
            "Kgs": b.COL_JS_KGS, "$ Prioridad": b.COL_JS_PRIORIDAD,
        })
        guias = b.build_jetsmart_guias(df)
        self.assertEqual(list(guias["Nro guía"]), [100186, 102141, 102177])
        self.assertTrue(all(isinstance(v, int) for v in guias["Nro guía"]))


class SinDatoSeMuestraGuion(_Historial):
    """Punto 2: nada de "None" en la tabla de excepciones."""

    def test_kg_e_ingreso_de_ariel_sin_dato_muestran_guion(self):
        at = self._abrir()
        exc = next(df.value for df in at.dataframe if "Tipo" in df.value.columns)
        pendiente = exc[exc["Tipo"] == EXC_PENDIENTE].iloc[0]
        self.assertEqual(pendiente["Kg Ariel"], "—")
        self.assertEqual(pendiente["Ingreso Ariel"], "—")
        self.assertEqual(pendiente["Fecha Ariel"], "—")
        # y los que si tienen dato siguen mostrando el numero
        dif = exc[exc["Tipo"] == EXC_DIFERENCIA].iloc[0]
        self.assertEqual(dif["Kg Ariel"], "6,00")
        self.assertEqual(dif["Ingreso Ariel"], "120,00")

    def test_no_aparece_none_ni_nan_en_ninguna_tabla(self):
        at = self._abrir()
        for df in at.dataframe:
            for v in df.value.astype(str).to_numpy().ravel():
                self.assertNotIn(v, ("None", "nan", "NaN", "<NA>"))


class MismosComponentesQueElResultado(_Historial):
    """Punto 3: tarjetas, "Para revisar" y detalle en expanders cerrados."""

    def test_tarjetas_arriba_con_los_numeros_clave(self):
        at = self._abrir()
        tarjetas = [h for h in _htmls(at) if "hwc-numcard" in h and "<style" not in h]
        etiquetas = " ".join(tarjetas)
        for texto in ("Total a entregar a WCS", "Guías liquidadas", "Para revisar"):
            self.assertIn(texto, etiquetas)
        self.assertTrue(any("CVLP, total final con IVA" in h for h in _htmls(at)))
        self.assertTrue(any(">3<" in h for h in tarjetas), "debe contar las 3 guias liquidadas")

    def test_bloque_para_revisar_con_las_guias_enteras(self):
        at = self._abrir()
        bloque = [h for h in _htmls(at) if 'class="hwc-review"' in h]
        self.assertEqual(len(bloque), 1)
        self.assertIn("Para revisar: 2 guías", bloque[0])
        celdas = re.findall(r'<td class="hwc-num">([^<]*)</td>', bloque[0])
        self.assertEqual(sorted(celdas), ["100186", "102177"])

    def test_sin_excepciones_no_hay_bloque_para_revisar(self):
        at = self._abrir(cruce=None)
        self.assertFalse(any('class="hwc-review"' in h for h in _htmls(at)))

    def test_la_vista_vieja_ya_no_esta(self):
        at = self._abrir()
        self.assertFalse(any("Excepciones del cruce" in h for h in _htmls(at) if "<style" not in h))

    def test_detalle_en_expanders_cerrados(self):
        at = self._abrir()
        exp = {e.label: e for e in at.get("status") if e.label.startswith(("Cómo se armó", "Liquidación:", "Tabla de guías"))}
        self.assertEqual(len(exp), 3, list(exp))
        for e in exp.values():
            self.assertFalse(e.proto.expanded, e.label)


class DatosUsadosConDesglose(_Historial):
    """La linea "Datos usados" del Historial usa el mismo desglose que el paso 3 y el resultado."""

    BASE = {"tipo_cambio": 1485.0, "vuelos_inter": {"EZE": 49, "MDZ": 27}, "usd_por_vuelo": {"EZE": 17.5, "MDZ": 35.0}}

    def _linea(self, at):
        return next(c.value for c in at.caption if c.value.startswith("Datos usados:"))

    def test_con_el_desglose_guardado_muestra_expo_impo_y_mdz(self):
        parametros = {**self.BASE, "vuelos_desglose": {"EZE": [["Expo", 25], ["Impo", 24]], "MDZ": []}}
        linea = self._linea(self._abrir(parametros=parametros))
        self.assertEqual(
            linea,
            "Datos usados: tipo de cambio $ 1.485,00 · EZE: Vuelos Expo 437,5 USD + Vuelos Impo 420 USD = 857,5 USD"
            " · MDZ: 945 USD.",
        )
        self.assertNotIn("49", linea)

    def test_la_estacion_cambiada_a_mano_queda_con_la_linea_simple(self):
        parametros = {**self.BASE, "vuelos_desglose": {"EZE": [["Expo", 25], ["Impo", 24]]}}  # MDZ no figura
        linea = self._linea(self._abrir(parametros=parametros))
        self.assertIn("EZE: Vuelos Expo 437,5 USD + Vuelos Impo 420 USD = 857,5 USD", linea)
        self.assertIn("MDZ 27 vuelos × 35 USD = 945 USD", linea)

    def test_una_liquidacion_vieja_sin_desglose_queda_como_siempre(self):
        linea = self._linea(self._abrir(parametros=self.BASE))
        self.assertEqual(linea, "Datos usados: tipo de cambio $ 1.485,00.")

    def test_si_el_desglose_no_suma_la_cantidad_guardada_no_se_usa(self):
        parametros = {**self.BASE, "vuelos_desglose": {"EZE": [["Expo", 25], ["Impo", 20]]}}
        linea = self._linea(self._abrir(parametros=parametros))
        self.assertIn("EZE 49 vuelos × 17,5 USD = 857,5 USD", linea)
        self.assertNotIn("Vuelos Expo", linea)
