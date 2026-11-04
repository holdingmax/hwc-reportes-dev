"""El mes anterior de JetSmart leido de la base (src/jetsmart_mes_anterior.py,
src/db.py, armar_conjunto y la pantalla).

Correr con:
    python -m unittest tests.test_jetsmart_mes_anterior -v

No necesita Postgres: el acceso a la base se prueba con una conexion
simulada que registra los SQL y los parametros (la migracion real se prueba
al aplicarla en Testing). Los tests con archivos reales (septiembre y
octubre 2026) usan tests/fixtures/jetsmart/, no versionado, y se saltean si
faltan.
"""

import os
import unittest
from datetime import date, datetime, timezone
from pathlib import Path
from unittest import mock

import pandas as pd
from streamlit.testing.v1 import AppTest

import src.db as db
from src.config import COL_JS_GUIA
from src.jetsmart_archivos import (
    TIPO_ARIEL,
    TIPO_GUIAS,
    ArchivoDetectado,
    armar_conjunto,
    cargar_guias_sistema,
)
from src.jetsmart_cruce import (
    COL_AR_GUIA,
    RES_EXCLUIDA,
    RES_INCLUIDA,
    RES_PENDIENTE,
    cruzar_jetsmart,
    leer_ariel,
)
from src.jetsmart_mes_anterior import (
    FUENTE_ARIEL,
    FUENTE_SISTEMA,
    LecturaMesAnterior,
    ResumenMesAnterior,
    fecha_corta,
    filas_guias_periodo,
    frames_desde_filas,
)

OCT, SEP = ("OCT", "26"), ("SEP", "26")
FIX = Path(__file__).resolve().parent / "fixtures" / "jetsmart"
G_OCT, G_SEP = FIX / "Guias_Aerea_10-2026.xlsx", FIX / "Guias_Aerea_09-2026.xlsx"
A_OCT, A_SEP = FIX / "Ariel_Rendicion_octubre_2026.xlsx", FIX / "Ariel_Rendicion_septiembre_2026.xlsx"
IVA = FIX / "IVA_VENTAS_HW_10-2026.xlsx"
hay_archivos = unittest.skipUnless(all(p.exists() for p in (G_OCT, G_SEP, A_OCT, A_SEP)), "faltan los archivos reales")

CARGADO = datetime(2026, 11, 4, 15, 30, tzinfo=timezone.utc)


def _resumen(version=2, periodo=SEP, origen="carga_inicial"):
    return ResumenMesAnterior(periodo, version, CARGADO, origen, 1984, 1978)


# ---------------------------------------------------------------------------
# Texto que ve la usuaria
# ---------------------------------------------------------------------------
class TextoDelMesAnterior(unittest.TestCase):
    def test_texto_con_version_y_fecha(self):
        self.assertEqual(_resumen().texto, "Mes anterior: septiembre 2026, versión 2, cargado el 4/11")

    def test_la_fecha_es_la_argentina(self):
        # 01:30 UTC del 5/11 son las 22:30 del 4/11 en Argentina (UTC-3, sin horario de verano).
        self.assertEqual(fecha_corta(datetime(2026, 11, 5, 1, 30, tzinfo=timezone.utc)), "4/11")
        self.assertEqual(fecha_corta(datetime(2026, 12, 1, 2, 0, tzinfo=timezone.utc)), "30/11")


# ---------------------------------------------------------------------------
# Armado de filas y reconstruccion (puro)
# ---------------------------------------------------------------------------
@hay_archivos
class FilasDelMes(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.g_oct, cls.g_sep = cargar_guias_sistema(G_OCT), cargar_guias_sistema(G_SEP)
        cls.a_oct, cls.a_sep = leer_ariel(A_OCT, cruda=True), leer_ariel(A_SEP, cruda=True)
        cls.cruce = cruzar_jetsmart(cls.g_oct, cls.a_oct.guias, OCT, cls.g_sep, cls.a_sep.guias,
                                    ariel_repetidas_identicas=cls.a_oct.repetidas_identicas)
        cls.filas_oct = filas_guias_periodo(cls.g_oct, cls.a_oct.guias, cls.cruce.destinos)
        cls.filas_sep = filas_guias_periodo(cls.g_sep, cls.a_sep.guias, None)

    def test_una_fila_por_guia_y_fuente_con_todas_las_guias(self):
        sistema = [f for f in self.filas_oct if f["fuente"] == FUENTE_SISTEMA]
        ariel = [f for f in self.filas_oct if f["fuente"] == FUENTE_ARIEL]
        self.assertEqual((len(sistema), len(ariel)), (1944, 1953))
        self.assertEqual(len({f["nro_guia"] for f in sistema}), 1944)   # sin repetidas
        self.assertEqual(len({f["nro_guia"] for f in ariel}), 1953)

    def test_resultado_de_cada_guia_en_octubre(self):
        por_guia = {}
        for f in self.filas_oct:
            por_guia.setdefault(f["nro_guia"], set()).add(f["resultado"])
        # las dos filas de una guia dicen lo mismo
        self.assertTrue(all(len(v) == 1 for v in por_guia.values()))
        cuenta = pd.Series([next(iter(v)) for v in por_guia.values()]).value_counts().to_dict()
        # 1.953 incluidas (1.934 match + 19 pendientes del mes anterior); 10 pendientes de declarar
        self.assertEqual(cuenta, {RES_INCLUIDA: 1953, RES_PENDIENTE: 10})

    def test_las_pendientes_estan_guardadas_y_son_solo_del_sistema(self):
        pend = [f for f in self.filas_oct if f["resultado"] == RES_PENDIENTE]
        self.assertEqual(len(pend), 10)
        self.assertEqual({f["fuente"] for f in pend}, {FUENTE_SISTEMA})
        self.assertEqual({f["motivo"] for f in pend}, {"Pendiente de declarar"})

    def test_carga_inicial_no_inventa_resultado(self):
        self.assertTrue(all(f["resultado"] is None and f["motivo"] is None for f in self.filas_sep))
        self.assertEqual(len([f for f in self.filas_sep if f["fuente"] == FUENTE_SISTEMA]), 1984)

    def test_tipos_simples_para_la_base(self):
        f = self.filas_oct[0]
        self.assertIsInstance(f["nro_guia"], int)
        self.assertIsInstance(f["fecha"], date)
        self.assertIsInstance(f["kg"], float)

    def test_el_mes_anterior_reconstruido_da_el_mismo_cruce(self):
        # Lo importante: leer septiembre de la base no cambia nada respecto de subir sus archivos.
        export_ant, ariel_ant = frames_desde_filas(self.filas_sep)
        r = cruzar_jetsmart(self.g_oct, self.a_oct.guias, OCT, export_ant, ariel_ant,
                            ariel_repetidas_identicas=self.a_oct.repetidas_identicas)
        self.assertEqual(len(r.incluidas), 1953)
        self.assertEqual(r.resumen["motivos_inclusion"], self.cruce.resumen["motivos_inclusion"])
        self.assertEqual(r.resumen["excepciones_por_tipo"], {"Pendiente de declarar": 10})
        pd.testing.assert_frame_equal(r.incluidas, self.cruce.incluidas)

    def test_frames_con_la_forma_que_espera_el_cruce(self):
        export_ant, ariel_ant = frames_desde_filas(self.filas_sep)
        self.assertEqual((len(export_ant), len(ariel_ant)), (1984, 1978))
        self.assertIn(COL_JS_GUIA, export_ant.columns)
        self.assertIn(COL_AR_GUIA, ariel_ant.columns)

    def test_frames_de_un_mes_vacio(self):
        export_ant, ariel_ant = frames_desde_filas([])
        r = cruzar_jetsmart(self.g_oct, self.a_oct.guias, OCT, export_ant, ariel_ant)
        self.assertGreater(len(r.incluidas), 0)


class DestinosPorRegla(unittest.TestCase):
    """Cada regla del cruce deja el destino correcto en la base (sin archivos reales)."""

    def _run(self, export_guias, ariel_guias, prev_ariel=(), prev_export=()):
        from tests.test_jetsmart_duplicados import _export_con
        from src.jetsmart_cruce import _AR_COLUMNS
        fila = [827, 1, 0, pd.Timestamp(2026, 10, 15), 3039, "C", "D", "AEP", "BRC", 1, 8, 8, 736, 10000, "X"]

        def ariel(guias):
            return pd.DataFrame([fila[:2] + [g] + fila[3:] for g in guias], columns=_AR_COLUMNS).assign(**{COL_AR_GUIA: lambda d: d[COL_AR_GUIA].astype("int64")})

        exp = pd.concat([_export_con(g) for g in export_guias]) if export_guias else _export_con(1).iloc[0:0]
        pa = ariel(prev_ariel) if prev_ariel else None
        pe = pd.concat([_export_con(g) for g in prev_export]) if prev_export else None
        return cruzar_jetsmart(exp, ariel(ariel_guias), OCT, pe, pa).destinos

    def test_destinos(self):
        d = self._run(export_guias=[10, 11, 12], ariel_guias=[10, 12, 20], prev_ariel=[12], prev_export=[20])
        self.assertEqual(d[10][0], RES_INCLUIDA)          # match
        self.assertEqual(d[11], (RES_PENDIENTE, "Pendiente de declarar"))   # solo en el sistema
        self.assertEqual(d[12][0], RES_EXCLUIDA)          # ya declarada el mes anterior
        self.assertEqual(d[20][0], RES_INCLUIDA)          # pendiente del mes anterior, declarada ahora
        self.assertEqual(set(d), {10, 11, 12, 20})


# ---------------------------------------------------------------------------
# armar_conjunto con el mes anterior en la base
# ---------------------------------------------------------------------------
def _ar(nombre, tipo, periodo):
    return ArchivoDetectado(nombre, tipo, cantidad=10, periodo=periodo)


def _lector(estado, version=2):
    return lambda periodo: LecturaMesAnterior(estado, _resumen(version, periodo) if estado == "ok" else None)


class ConjuntoConMesAnteriorGuardado(unittest.TestCase):
    OCTUBRE = [_ar("g_oct", TIPO_GUIAS, OCT), _ar("a_oct", TIPO_ARIEL, OCT)]

    def test_con_el_mes_guardado_no_hay_que_subir_nada_mas(self):
        c = armar_conjunto(self.OCTUBRE, _lector("ok"))
        self.assertTrue(c.listo)
        self.assertEqual(c.origen_mes_anterior, "bd")
        self.assertEqual(c.mensaje_mes_anterior, "Mes anterior: septiembre 2026, versión 2, cargado el 4/11")

    def test_carga_inicial_cuando_la_base_no_lo_tiene(self):
        c = armar_conjunto(self.OCTUBRE, _lector("sin_datos"))
        self.assertFalse(c.listo)
        self.assertEqual(len(c.faltantes), 1)
        self.assertIn("Todavía no tengo guardado septiembre 2026", c.faltantes[0])
        self.assertIn("una sola vez", c.faltantes[0])

    def test_carga_inicial_con_los_archivos_del_mes_anterior(self):
        archivos = self.OCTUBRE + [_ar("g_sep", TIPO_GUIAS, SEP), _ar("a_sep", TIPO_ARIEL, SEP)]
        c = armar_conjunto(archivos, _lector("sin_datos"))
        self.assertTrue(c.listo)
        self.assertEqual(c.origen_mes_anterior, "archivos")
        self.assertIn("se guarda", c.mensaje_mes_anterior)

    def test_sin_base_pide_los_archivos_y_avisa(self):
        c = armar_conjunto(self.OCTUBRE, _lector("sin_base"))
        self.assertFalse(c.listo)
        self.assertTrue(any("No pude consultar la base" in a for a in c.avisos))
        self.assertTrue(all("Todavía no tengo guardado" not in f for f in c.faltantes))
        self.assertEqual(len(c.faltantes), 2)

    def test_sin_lector_se_comporta_como_antes(self):
        c = armar_conjunto(self.OCTUBRE)
        self.assertEqual(len(c.faltantes), 2)
        self.assertIsNone(c.origen_mes_anterior)

    def test_los_dos_archivos_del_mes_anterior_reemplazan_la_version_guardada(self):
        archivos = self.OCTUBRE + [_ar("g_sep", TIPO_GUIAS, SEP), _ar("a_sep", TIPO_ARIEL, SEP)]
        c = armar_conjunto(archivos, _lector("ok", version=3))
        self.assertEqual(c.origen_mes_anterior, "archivos")
        self.assertTrue(any("en lugar de la versión 3 guardada" in a for a in c.avisos))

    def test_un_solo_archivo_del_mes_anterior_no_mezcla_con_lo_guardado(self):
        c = armar_conjunto(self.OCTUBRE + [_ar("g_sep", TIPO_GUIAS, SEP)], _lector("ok", version=3))
        self.assertEqual(c.origen_mes_anterior, "bd")
        self.assertIsNone(c.export_anterior)
        self.assertTrue(any("No usé g_sep" in a and "versión 3" in a for a in c.avisos))

    def test_sin_archivos_no_consulta_la_base(self):
        lector = mock.Mock()
        armar_conjunto([], lector)
        lector.assert_not_called()


# ---------------------------------------------------------------------------
# db.py con una conexion simulada
# ---------------------------------------------------------------------------
class _Transaccion:
    def __init__(self, log, nombre):
        self.log, self.nombre = log, nombre

    def __enter__(self):
        self.log.append(("begin", self.nombre))
        return self

    def __exit__(self, tipo, valor, tb):
        self.log.append(("end", self.nombre, tipo.__name__ if tipo else None))
        return False


class _Cursor:
    def __init__(self, conn):
        self.conn, self.sql = conn, ""

    def __enter__(self):
        return self

    def __exit__(self, *a):
        return False

    def execute(self, sql, params=None):
        self.sql = " ".join(sql.split())
        self.conn.log.append(("execute", self.sql, params))

    def executemany(self, sql, filas):
        filas = list(filas)
        self.conn.log.append(("executemany", " ".join(sql.split()), filas))
        if self.conn.falla_executemany:
            raise RuntimeError("tabla inexistente")

    def _respuesta(self, indice):
        for clave, uno, varios in self.conn.respuestas:
            if clave in self.sql:
                return (uno, varios)[indice]
        return None

    def fetchone(self):
        return self._respuesta(0)

    def fetchall(self):
        return self._respuesta(1)


class _Conexion:
    def __init__(self, respuestas, falla_executemany=False):
        self.log, self.respuestas, self.falla_executemany = [], respuestas, falla_executemany
        self._n = 0

    def transaction(self):
        self._n += 1
        return _Transaccion(self.log, f"tx{self._n}")

    def cursor(self):
        return _Cursor(self)


def _con_conexion(conn):
    return mock.patch.object(db, "_get_connection", mock.MagicMock(return_value=conn))


RESPUESTAS_GUARDAR = [
    ("SELECT id FROM cargas_archivo", None, None),
    ("INSERT INTO cargas_archivo", (10,), None),
    ("INSERT INTO liquidaciones", (55,), None),
    ("COALESCE(MAX(version)", (3,), None),
    ("INSERT INTO jetsmart_periodos", (77,), None),
]


def _guardar(conn, meses):
    guias = pd.DataFrame({"Nro guía": [1, 2], "TOTAL": [100.0, 200.0]})
    with _con_conexion(conn):
        return db.guardar_liquidacion_jetsmart(
            "g.xlsx", b"abc", OCT, 2, 0, guias, {"total_wcs": 300.0}, {"tipo_cambio": 1485.0}, None, meses,
        )


def _mes(origen, n=2, sin_resultado=False):
    """Mes de ejemplo para guardar. Una 'liquidacion' trae resultado en cada guia (es la regla);
    la 'carga_inicial' no. sin_resultado=True arma el caso invalido para probar que se rechaza."""
    con_resultado = origen == "liquidacion" and not sin_resultado
    filas = [
        {"fuente": FUENTE_SISTEMA, "nro_guia": i, "fecha": date(2026, 10, 1), "origen": "AEP", "kg": 1.5,
         "ingreso": 100.0, "estado": "Entregada",
         "resultado": "incluida" if con_resultado else None, "motivo": "match" if con_resultado else None}
        for i in range(n)
    ]
    return {"period": SEP if origen == "carga_inicial" else OCT, "origen_dato": origen, "filas": filas,
            "archivos": [{"nombre": "x.xlsx", "hash": "h", "tipo": "sistema"}], "repetidas": 20}


class GuardadoEnLaBase(unittest.TestCase):
    def test_guarda_mes_con_version_liquidacion_y_guias(self):
        conn = _Conexion(RESPUESTAS_GUARDAR)
        resultado = _guardar(conn, [_mes("carga_inicial", 3), _mes("liquidacion", 2)])
        self.assertEqual(resultado, {"liquidacion": True, "meses": True})
        periodos = [e for e in conn.log if e[0] == "execute" and "INSERT INTO jetsmart_periodos" in e[1]]
        self.assertEqual(len(periodos), 2)
        carga, liq = periodos[0][2], periodos[1][2]
        # (periodo, version, origen_dato, liquidacion_id, archivos, repetidas)
        self.assertEqual(carga[:4], (date(2026, 9, 1), 3, "carga_inicial", None))   # la carga inicial no tiene liquidacion
        self.assertEqual(liq[:4], (date(2026, 10, 1), 3, "liquidacion", 55))        # la liquidacion si
        self.assertEqual(liq[5], 20)
        lotes = [e for e in conn.log if e[0] == "executemany"]
        self.assertEqual([len(l[2]) for l in lotes], [3, 2])
        self.assertEqual(lotes[0][2][0][0], 77)   # periodo_id

    def test_la_version_siguiente_se_calcula_en_la_base(self):
        conn = _Conexion(RESPUESTAS_GUARDAR)
        _guardar(conn, [_mes("liquidacion")])
        consulta = next(e for e in conn.log if e[0] == "execute" and "COALESCE(MAX(version)" in e[1])
        self.assertIn("FROM jetsmart_periodos WHERE periodo = %s", consulta[1])
        self.assertEqual(consulta[2], (date(2026, 10, 1),))

    def test_si_falla_el_guardado_de_las_guias_la_liquidacion_se_guarda_igual(self):
        conn = _Conexion(RESPUESTAS_GUARDAR, falla_executemany=True)
        resultado = _guardar(conn, [_mes("liquidacion")])
        self.assertEqual(resultado, {"liquidacion": True, "meses": False})
        # el fallo quedo contenido en el savepoint: la transaccion externa termino sin error
        cierres = [e for e in conn.log if e[0] == "end"]
        self.assertEqual(cierres[0][2], "RuntimeError")      # savepoint (tx2) abortado...
        self.assertIsNone(cierres[1][2])                       # ...y la externa (tx1) confirmada

    def test_sin_base_no_hace_nada_y_no_lanza(self):
        with mock.patch.object(db, "_get_connection", mock.MagicMock(return_value=None)):
            r = db.guardar_liquidacion_jetsmart("g", b"x", OCT, 1, 0, pd.DataFrame({"a": [1]}), {}, {}, None, [_mes("liquidacion")])
        self.assertEqual(r, {"liquidacion": False, "meses": False})

    def test_sin_meses_equivale_a_antes(self):
        conn = _Conexion(RESPUESTAS_GUARDAR)
        self.assertEqual(_guardar(conn, None), {"liquidacion": True, "meses": True})
        self.assertFalse([e for e in conn.log if "jetsmart_periodos" in str(e)])


class LecturaDeLaBase(unittest.TestCase):
    FILA_VIGENTE = (9, 2, CARGADO, "carga_inicial")
    COLUMNAS = [
        (FUENTE_SISTEMA, 101, date(2026, 9, 1), "AEP", __import__("decimal").Decimal("1.500"), __import__("decimal").Decimal("12000.0000"), "Entregada", None, None),
        (FUENTE_ARIEL, 101, date(2026, 9, 1), "AEP", __import__("decimal").Decimal("1.500"), __import__("decimal").Decimal("12000.0000"), None, None, None),
    ]

    def _conn(self, vigente=FILA_VIGENTE):
        return _Conexion([
            ("FROM jetsmart_periodo_vigente", vigente, None),
            ("GROUP BY fuente", None, [("sistema", 1984), ("ariel", 1978)]),
            ("SELECT fuente, nro_guia", None, self.COLUMNAS),
        ])

    def test_resumen_con_version_fecha_y_cantidades(self):
        with _con_conexion(self._conn()):
            lectura = db.obtener_mes_anterior_jetsmart(SEP)
        self.assertEqual(lectura.estado, "ok")
        self.assertEqual((lectura.resumen.version, lectura.resumen.n_sistema, lectura.resumen.n_ariel), (2, 1984, 1978))
        self.assertEqual(lectura.resumen.texto, "Mes anterior: septiembre 2026, versión 2, cargado el 4/11")
        self.assertEqual(lectura.filas, [])   # sin pedir las guias, no se traen

    def test_con_guias_convierte_decimales_a_float(self):
        with _con_conexion(self._conn()):
            lectura = db.obtener_mes_anterior_jetsmart(SEP, con_guias=True)
        self.assertEqual(len(lectura.filas), 2)
        self.assertIsInstance(lectura.filas[0]["kg"], float)
        self.assertEqual(lectura.filas[0]["ingreso"], 12000.0)

    def test_lee_la_vista_de_versiones_vigentes_del_mes_pedido(self):
        conn = self._conn()
        with _con_conexion(conn):
            db.obtener_mes_anterior_jetsmart(SEP)
        consulta = next(e for e in conn.log if "jetsmart_periodo_vigente" in e[1])
        self.assertEqual(consulta[2], (date(2026, 9, 1),))

    def test_mes_que_no_esta(self):
        with _con_conexion(self._conn(vigente=None)):
            self.assertEqual(db.obtener_mes_anterior_jetsmart(SEP).estado, "sin_datos")

    def test_sin_base(self):
        with mock.patch.object(db, "_get_connection", mock.MagicMock(return_value=None)):
            self.assertEqual(db.obtener_mes_anterior_jetsmart(SEP).estado, "sin_base")

    def test_error_de_la_base_no_lanza(self):
        malo = mock.MagicMock()
        malo.cursor.side_effect = RuntimeError("caida")
        with mock.patch.object(db, "_get_connection", mock.MagicMock(return_value=malo)):
            self.assertEqual(db.obtener_mes_anterior_jetsmart(SEP).estado, "sin_base")


# ---------------------------------------------------------------------------
# Pantalla: el mes anterior se lee de la base y se muestra con su version
# ---------------------------------------------------------------------------
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
APP = Path(__file__).resolve().parent.parent / "app.py"


class NroGuiaEntero(unittest.TestCase):
    def test_decimales_de_coma_flotante_vuelven_a_entero(self):
        from src.jetsmart_cruce import nro_guia_entero
        r = nro_guia_entero(pd.Series([102176.99999999999, 102177.0, 5, None, "x"]))
        self.assertEqual(r.tolist()[:3], [102177, 102177, 5])
        self.assertTrue(r.isna().tolist()[3:] == [True, True])
        self.assertTrue(pd.api.types.is_integer_dtype(r))


@hay_archivos
class PantallaConMesAnteriorGuardado(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        g_sep, a_sep = cargar_guias_sistema(G_SEP), leer_ariel(A_SEP, cruda=True)
        cls.filas_sep = filas_guias_periodo(g_sep, a_sep.guias, None)

    def setUp(self):
        self._env = mock.patch.dict(os.environ)
        self._env.start()
        os.environ.pop("DATABASE_URL", None)
        self.guardar = mock.patch("src.db.guardar_liquidacion_jetsmart", return_value={"liquidacion": True, "meses": True}).start()
        self.leer = mock.patch("src.db.obtener_mes_anterior_jetsmart").start()
        self.addCleanup(mock.patch.stopall)
        self.addCleanup(self._env.stop)

    def _lectura(self, estado, con_guias=False):
        if estado == "ok":
            return LecturaMesAnterior("ok", _resumen(2), self.filas_sep if con_guias else [])
        return LecturaMesAnterior(estado)

    def _app(self, archivos, estado="ok"):
        self.leer.side_effect = lambda periodo, con_guias=False: self._lectura(estado, con_guias)
        at = AppTest.from_file(str(APP), default_timeout=180)
        at.session_state["aerolinea_sel"] = "jetsmart"
        at.run()
        at.get("file_uploader")[0].set_value([(p.name, p.read_bytes(), XLSX) for p in archivos])
        at.run()
        return at

    @staticmethod
    def _textos(at):
        return "\n".join(
            [m.value for m in at.markdown] + [s.value for s in at.success]
            + [w.value for w in at.warning] + [e.value for e in at.error]
        )

    def test_aviso_verde_si_el_mes_anterior_ya_esta_guardado(self):
        at = self._app([G_OCT, A_OCT])
        verdes = [s.value for s in at.success]
        self.assertEqual(len(verdes), 1)
        self.assertIn("Septiembre 2026 ya está guardado, no hace falta subirlo", verdes[0])
        self.assertIn("versión 2, cargado el 4/11", verdes[0])
        self.assertNotIn("todavía no está guardado", self._textos(at))

    def test_aviso_amarillo_si_falta_el_mes_anterior(self):
        at = self._app([G_OCT, A_OCT], estado="sin_datos")
        self.assertEqual([s.value for s in at.success], [])
        amarillos = [w.value for w in at.warning if "todavía no está guardado" in w.value]
        self.assertEqual(len(amarillos), 1)
        self.assertIn("Septiembre 2026 todavía no está guardado", amarillos[0])
        self.assertIn("las guías de septiembre 2026 y la rendición de Ariel de septiembre 2026", amarillos[0])
        self.assertIn("Es una sola vez", amarillos[0])

    def test_muestra_version_y_fecha_y_no_pide_septiembre(self):
        at = self._app([G_OCT, A_OCT])
        texto = self._textos(at)
        self.assertIn("versión 2, cargado el 4/11", texto)
        self.assertNotIn("Falta", texto)
        at.number_input[0].set_value(1485.0)
        at.run()
        self.assertFalse(at.button[0].disabled)

    def test_procesa_octubre_con_septiembre_de_la_base(self):
        at = self._app([G_OCT, A_OCT])
        at.number_input[0].set_value(1485.0)
        at.run()
        at.button[0].click()
        at.run()
        self.assertEqual(len(at.exception), 0)
        self.assertEqual([e.value for e in at.error], [])
        kwargs = self.guardar.call_args.kwargs
        # Mismo resultado validado por Anita que subiendo los archivos de septiembre.
        self.assertEqual(len(kwargs["guias"]), 1953)
        self.assertIn("Mes anterior: septiembre 2026, versión 2, cargado el 4/11", self._textos(at))
        # Solo se guarda octubre (septiembre ya estaba): una version de la liquidacion con TODAS las guias.
        meses = kwargs["meses"]
        self.assertEqual([m["origen_dato"] for m in meses], ["liquidacion"])
        self.assertEqual(meses[0]["period"], OCT)
        self.assertEqual(len(meses[0]["filas"]), 1944 + 1953)
        resultados = {f["resultado"] for f in meses[0]["filas"]}
        self.assertEqual(resultados, {"incluida", "pendiente"})
        self.assertEqual(meses[0]["repetidas"], 20)
        self.assertEqual({a["tipo"] for a in meses[0]["archivos"]}, {"sistema", "ariel"})

    def test_carga_inicial_guarda_septiembre_y_octubre(self):
        at = self._app([G_OCT, A_OCT, G_SEP, A_SEP], estado="sin_datos")
        self.assertIn("se guarda", self._textos(at))
        at.number_input[0].set_value(1485.0)
        at.run()
        at.button[0].click()
        at.run()
        self.assertEqual(len(at.exception), 0)
        meses = self.guardar.call_args.kwargs["meses"]
        self.assertEqual([(m["origen_dato"], m["period"]) for m in meses], [("carga_inicial", SEP), ("liquidacion", OCT)])
        self.assertTrue(all(f["resultado"] is None for f in meses[0]["filas"]))   # septiembre no se cruzo aca
        self.assertEqual(len(self.guardar.call_args.kwargs["guias"]), 1953)

    def test_carga_inicial_pide_septiembre_una_sola_vez(self):
        at = self._app([G_OCT, A_OCT], estado="sin_datos")
        texto = self._textos(at)
        self.assertIn("Todavía no tengo guardado septiembre 2026", texto)
        self.assertIn("una sola vez", texto)
        self.assertTrue(at.button[0].disabled)

    def test_tabla_de_excepciones_muestra_el_nro_de_guia_entero(self):
        at = self._app([G_OCT, A_OCT])
        at.number_input[0].set_value(1485.0)
        at.run()
        at.button[0].click()
        at.run()
        tablas = [d.value for d in at.dataframe if "Nro guía" in d.value.columns and "Tipo" in d.value.columns]
        self.assertTrue(tablas, "no se mostro la tabla de excepciones")
        nros = tablas[0]["Nro guía"]
        self.assertTrue(pd.api.types.is_integer_dtype(nros), f"dtype {nros.dtype}")
        self.assertTrue(all(float(n).is_integer() for n in nros))

    def test_si_no_se_pudieron_guardar_las_guias_del_mes_avisa(self):
        self.guardar.return_value = {"liquidacion": True, "meses": False}
        at = self._app([G_OCT, A_OCT])
        at.number_input[0].set_value(1485.0)
        at.run()
        at.button[0].click()
        at.run()
        self.assertIn("no pude guardar las guías", self._textos(at))

    def test_si_el_mes_guardado_deja_de_estar_disponible_al_procesar_lo_explica(self):
        at = self._app([G_OCT, A_OCT])
        at.number_input[0].set_value(1485.0)
        at.run()
        # El resumen se lee bien, pero la lectura completa de las guias falla al procesar.
        self.leer.side_effect = lambda periodo, con_guias=False: (
            LecturaMesAnterior("sin_base") if con_guias else self._lectura("ok"))
        at.button[0].click()
        at.run()
        errores = " ".join(e.value for e in at.error)
        self.assertIn("No pude leer del sistema el mes anterior guardado", errores)
        self.assertNotIn("MesAnteriorNoDisponible", errores)
        self.assertEqual(len(at.exception), 0)
        self.guardar.assert_not_called()

    def test_si_la_base_deja_de_responder_justo_al_hacer_clic_dice_que_falta(self):
        at = self._app([G_OCT, A_OCT])
        at.number_input[0].set_value(1485.0)
        at.run()
        # Al volver a correr la pantalla con el clic, el mes anterior ya no se encuentra.
        self.leer.side_effect = lambda periodo, con_guias=False: LecturaMesAnterior("sin_base")
        at.button[0].click()
        at.run()
        errores = " ".join(e.value for e in at.error)
        self.assertIn("Todavía no se puede procesar", errores)
        self.assertIn("mes anterior", errores)
        self.assertNotIn("ValueError", errores)
        self.assertEqual(len(at.exception), 0)
        self.guardar.assert_not_called()


if __name__ == "__main__":
    unittest.main()
