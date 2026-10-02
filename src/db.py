"""Persistencia en Postgres de cargas, movimientos y liquidaciones (etapa 1).

Capa aditiva pura: no cambia ni depende de ningun calculo de negocio. No
importa nada de loader.py/report_builder.py/liquidacion_builder.py/
period_utils.py mas alla de leer AIRLINE_CONFIGS/CHARGE_TYPES/COL_* de
config.py (que son datos de configuracion, no calculo) para no duplicar
el mapeo de nombre-de-hoja -> tipo de cargo/estacion.

Las funciones publicas de guardado (guardar_reporte_simple,
guardar_liquidacion_latam, guardar_liquidacion_jetsmart) NUNCA lanzan excepciones: si DATABASE_URL no esta configurada, o la base no
responde, no hacen nada. La app tiene que poder seguir generando y
descargando el Excel exactamente igual aunque la base este caida -- la
persistencia es un efecto secundario, no un requisito del flujo de Cristian.

tarifas y ajustes_manuales se crean por migracion (ver sql/migrations/) pero
no se leen ni se escriben desde este modulo todavia: el calculo sigue
usando AVIANCA_DELIVERY_FEE_USD_POR_ESTACION de config.py como hasta ahora,
y "Compensacion" en LATAM sigue siendo el texto manual de siempre.

Ademas de las funciones de guardado, este modulo expone dos funciones de
SOLO LECTURA para la pantalla de Historial (obtener_filtros_historial /
obtener_historial): nunca escriben nada, y siguen el mismo criterio
fail-soft que el guardado -- si la base no responde, devuelven None en vez
de lanzar excepcion, para que la pantalla muestre un mensaje en vez de
romperse.
"""

import hashlib
import os
from datetime import date

import pandas as pd
import psycopg
import streamlit as st
from dotenv import load_dotenv
from psycopg.types.json import Json

from src.config import (
    AIRLINE_CONFIGS,
    CHARGE_TYPES,
    COL_ADUANA,
    COL_AEROLINEA,
    COL_CLIENTE,
    COL_CODIGO,
    COL_COD_VUELO,
    COL_COLLECT,
    COL_CONDICION,
    COL_DRY_FEE,
    COL_ESTACION,
    COL_IATA,
    COL_TIPO,
    COL_TPO_CAMBIO,
    COL_TRANS_E,
    LATAM_LIQUIDACION_STATION,
)
from src.period_utils import Period, detect_period, excluded_by_period

load_dotenv()

_MESES = {
    "ENE": 1, "FEB": 2, "MAR": 3, "ABR": 4, "MAY": 5, "JUN": 6,
    "JUL": 7, "AGO": 8, "SEP": 9, "OCT": 10, "NOV": 11, "DIC": 12,
}

_MOVIMIENTO_COLUMNS = [
    COL_CODIGO, COL_COD_VUELO, COL_CLIENTE, COL_CONDICION, COL_TIPO,
    COL_TPO_CAMBIO, COL_DRY_FEE, COL_ADUANA, COL_TRANS_E, COL_IATA,
    COL_COLLECT, COL_AEROLINEA, COL_ESTACION,
]


def _period_to_date(period: Period) -> date:
    mes, anio = period
    return date(2000 + int(anio), _MESES[mes], 1)


def _texto_o_none(value) -> str | None:
    return None if pd.isna(value) else str(value)


def _numero_o_none(value) -> float | None:
    return None if pd.isna(value) else float(value)


def _codigo_texto(value) -> str | None:
    """Igual criterio que _format_codigo en report_builder.py, copiado (no
    importado) para no acoplar este modulo a un simbolo privado de otro."""
    if pd.isna(value):
        return None
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def _jsonable(value):
    """Convierte numpy scalars (numpy.float64, etc.) a tipos nativos de
    Python antes de serializar a JSONB -- json.dumps no sabe con numpy."""
    if isinstance(value, dict):
        return {k: _jsonable(v) for k, v in value.items()}
    if isinstance(value, (list, tuple)):
        return [_jsonable(v) for v in value]
    if hasattr(value, "item"):
        return value.item()
    return value


@st.cache_resource
def _get_connection() -> psycopg.Connection | None:
    database_url = os.environ.get("DATABASE_URL")
    if not database_url:
        return None
    # connect_timeout corto: si la base esta inalcanzable (URL mal, firewall,
    # caida de red), esto tiene que fallar rapido, no colgar la app entera
    # esperando el timeout de TCP del sistema operativo (que puede ser de
    # varios minutos).
    #
    # autocommit=True es OBLIGATORIO aca, no cosmetico. Esta conexion vive
    # cacheada (@st.cache_resource) y se REUSA entre reruns de Streamlit --
    # y como st.tabs() ejecuta el codigo de TODOS los tabs en cada rerun
    # (no solo el visible), un mismo rerun de "Generar reporte" tambien
    # corre las lecturas de Historial (obtener_filtros_historial /
    # obtener_historial), que hacen cur.execute(SELECT) sueltos sin
    # with conn.transaction(). Sin autocommit, cada SELECT abre una
    # transaccion implicita que nadie cierra; el PROXIMO
    # with conn.transaction() (en guardar_reporte_simple) encuentra la
    # conexion ya "dentro" de una transaccion y crea un SAVEPOINT anidado
    # en vez de una transaccion nueva -- el INSERT se ve desde la MISMA
    # conexion (por eso la app mostraba el dato bien) pero nunca se
    # comittea de verdad, y desaparece en cuanto la conexion se recicla.
    # Confirmado reproduciendo el bug con una conexion nueva independiente
    # antes de este fix. Con autocommit=True cada statement suelto
    # comittea solo, y with conn.transaction() sigue agrupando varias
    # sentencias en una transaccion real (psycopg3 lo soporta igual en
    # modo autocommit).
    return psycopg.connect(database_url, connect_timeout=5, autocommit=True)


def _guardar_carga(conn: psycopg.Connection, nombre_archivo: str, file_bytes: bytes, df: pd.DataFrame) -> int:
    """Inserta cargas_archivo + movimientos_awb, o reusa la carga existente
    si este mismo archivo (por contenido) ya se habia subido antes."""
    archivo_hash = hashlib.sha256(file_bytes).hexdigest()

    with conn.cursor() as cur:
        cur.execute("SELECT id FROM cargas_archivo WHERE archivo_hash = %s", (archivo_hash,))
        existente = cur.fetchone()
        if existente:
            return existente[0]

        period = detect_period(df)
        excluded = excluded_by_period(df, period)
        periodo_mes, periodo_anio = period

        cur.execute(
            """
            INSERT INTO cargas_archivo
                (nombre_archivo, archivo_hash, periodo_mes, periodo_anio, periodo,
                 filas_totales, filas_excluidas_periodo)
            VALUES (%s, %s, %s, %s, %s, %s, %s)
            RETURNING id
            """,
            (nombre_archivo, archivo_hash, periodo_mes, periodo_anio,
             _period_to_date(period), len(df), len(excluded)),
        )
        carga_id = cur.fetchone()[0]
        _guardar_movimientos(cur, carga_id, df)
        return carga_id


def _guardar_movimientos(cur: psycopg.Cursor, carga_id: int, df: pd.DataFrame) -> None:
    registros = []
    for _, fila in df[_MOVIMIENTO_COLUMNS].iterrows():
        registros.append((
            carga_id,
            _codigo_texto(fila[COL_CODIGO]),
            _texto_o_none(fila[COL_COD_VUELO]),
            _texto_o_none(fila[COL_CLIENTE]),
            _texto_o_none(fila[COL_CONDICION]),
            _texto_o_none(fila[COL_TIPO]),
            _numero_o_none(fila[COL_TPO_CAMBIO]),
            _numero_o_none(fila[COL_DRY_FEE]),
            _numero_o_none(fila[COL_ADUANA]),
            _numero_o_none(fila[COL_TRANS_E]),
            _numero_o_none(fila[COL_IATA]),
            _numero_o_none(fila[COL_COLLECT]),
            _texto_o_none(fila[COL_AEROLINEA]),
            _texto_o_none(fila[COL_ESTACION]),
        ))
    if not registros:
        return
    cur.executemany(
        """
        INSERT INTO movimientos_awb
            (carga_id, codigo, cod_vuelo, cliente, condicion, tipo, tpo_cambio,
             dry_fee, aduana, trans_e, iata, collect, aerolinea, estacion)
        VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
        """,
        registros,
    )


def _guardar_liquidacion(
    conn: psycopg.Connection,
    carga_id: int,
    aerolinea: str,
    estacion: str,
    tipo_cargo: str,
    period: Period,
    cantidad_filas: int,
    monto_total: float | None,
    detalle_totales: dict | None = None,
) -> None:
    periodo_mes, periodo_anio = period
    with conn.cursor() as cur:
        cur.execute(
            """
            INSERT INTO liquidaciones
                (carga_id, aerolinea, estacion, tipo_cargo, periodo_mes, periodo_anio,
                 periodo, cantidad_filas, monto_total, detalle_totales)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            """,
            (
                carga_id, aerolinea, estacion, tipo_cargo, periodo_mes, periodo_anio,
                _period_to_date(period), cantidad_filas, monto_total,
                Json(_jsonable(detalle_totales)) if detalle_totales is not None else None,
            ),
        )


def _parse_sheet_name(sheet_name: str, airline_key: str) -> tuple[str, str]:
    """(tipo_cargo, estacion) a partir de un nombre de hoja generado.

    Reusa CHARGE_TYPES/AIRLINE_CONFIGS (config.py) en vez de duplicar el
    mapeo sheet -> tipo de cargo: si cambia un sheet_prefix alla, esto
    sigue funcionando sin tocar nada aca.
    """
    airline_cfg = AIRLINE_CONFIGS[airline_key]
    for charge_type_key in airline_cfg["charge_types"]:
        charge_cfg = CHARGE_TYPES[charge_type_key]
        if charge_cfg.get("per_station"):
            prefix = charge_cfg["sheet_prefix"] + " "
            if sheet_name.startswith(prefix):
                return charge_type_key, sheet_name[len(prefix):]
        elif sheet_name == charge_cfg.get("sheet_name"):
            return charge_type_key, "AR"  # hoja agrupada nacional (ej. "Collect AR")
    return "desconocido", "desconocido"  # salvaguarda, no deberia ocurrir


def _sumar_montos(sheet_df: pd.DataFrame, amount_columns: list[str]) -> float:
    total = 0.0
    for col in amount_columns:
        if col in sheet_df.columns:
            total += float(pd.to_numeric(sheet_df[col], errors="coerce").fillna(0).sum())
    return total


def guardar_reporte_simple(
    nombre_archivo: str,
    file_bytes: bytes,
    df: pd.DataFrame,
    airline_key: str,
    period: Period,
    sheets: dict[str, pd.DataFrame],
) -> None:
    """Persiste la carga completa (todas las aerolineas del archivo) mas una
    fila en liquidaciones por cada hoja generada para airline_key.

    No lanza excepciones (ver docstring del modulo): un fallo de base nunca
    debe poder bloquear la descarga del Excel que ya se genero.
    """
    try:
        conn = _get_connection()
        if conn is None:
            return
        with conn.transaction():
            carga_id = _guardar_carga(conn, nombre_archivo, file_bytes, df)
            for sheet_name, sheet_df in sheets.items():
                tipo_cargo, estacion = _parse_sheet_name(sheet_name, airline_key)
                amount_columns = CHARGE_TYPES.get(tipo_cargo, {}).get("amount_columns", [])
                monto_total = _sumar_montos(sheet_df, amount_columns)
                _guardar_liquidacion(
                    conn, carga_id, airline_key, estacion, tipo_cargo,
                    period, len(sheet_df), monto_total,
                )
    except Exception:
        _get_connection.clear()  # conexion posiblemente rota / conexion nunca se establecio: reintentar la proxima vez


def guardar_liquidacion_latam(
    nombre_archivo: str,
    file_bytes: bytes,
    df: pd.DataFrame,
    period: Period,
    detalle: pd.DataFrame,
    resumen: dict,
) -> None:
    """Persiste la carga completa mas una liquidacion 'latam_liquidacion'
    con el desglose de resumen (total_la/neto_gravado/iva/total_4m/
    total_periodo/sums) en detalle_totales.

    No lanza excepciones, igual que guardar_reporte_simple.
    """
    try:
        conn = _get_connection()
        if conn is None:
            return
        with conn.transaction():
            carga_id = _guardar_carga(conn, nombre_archivo, file_bytes, df)
            _guardar_liquidacion(
                conn, carga_id, "latam", LATAM_LIQUIDACION_STATION, "latam_liquidacion",
                period, len(detalle), resumen.get("total_periodo"), resumen,
            )
    except Exception:
        _get_connection.clear()


def guardar_liquidacion_jetsmart(
    nombre_archivo: str,
    file_bytes: bytes,
    period: Period,
    filas_export: int,
    filas_otros_meses: int,
    guias: pd.DataFrame,
    resumen: dict,
    parametros: dict,
    cruce: dict | None = None,
) -> None:
    """Persiste la carga del export de JetSmart mas una liquidacion
    'jetsmart_liquidacion' (estacion "AR": es una liquidacion nacional).

    A diferencia de Avianca/Gol/LATAM, NO escribe movimientos_awb: el
    export de JetSmart tiene otras columnas (KGs, $ Prioridad, sin
    Cod.Vuelo) que no entran en esa tabla sin una migracion. En su lugar,
    detalle_totales guarda el resumen, los datos manuales usados (TC,
    vuelos), las guias y, si la liquidacion se armo con el cruce contra
    Ariel, su resumen y excepciones (cruce), que es todo lo que hace falta
    para reconstruir el detalle y el Excel desde el historial (ver
    obtener_detalle_totales).

    monto_total = "Total collections to be delivered to WCS".

    No lanza excepciones, igual que el resto de las funciones de guardado.
    """
    try:
        conn = _get_connection()
        if conn is None:
            return
        archivo_hash = hashlib.sha256(file_bytes).hexdigest()
        periodo_mes, periodo_anio = period
        with conn.transaction():
            with conn.cursor() as cur:
                cur.execute("SELECT id FROM cargas_archivo WHERE archivo_hash = %s", (archivo_hash,))
                existente = cur.fetchone()
                if existente:
                    carga_id = existente[0]
                else:
                    cur.execute(
                        """
                        INSERT INTO cargas_archivo
                            (nombre_archivo, archivo_hash, periodo_mes, periodo_anio, periodo,
                             filas_totales, filas_excluidas_periodo)
                        VALUES (%s, %s, %s, %s, %s, %s, 0)
                        RETURNING id
                        """,
                        (nombre_archivo, archivo_hash, periodo_mes, periodo_anio,
                         _period_to_date(period), filas_export),
                    )
                    carga_id = cur.fetchone()[0]
            detalle_totales = {
                "resumen": resumen,
                "parametros": parametros,
                "filas_otros_meses": filas_otros_meses,
                "cruce": cruce,
                "guias": guias.astype(object).where(guias.notna(), None).to_dict(orient="records"),
            }
            _guardar_liquidacion(
                conn, carga_id, "jetsmart", "AR", "jetsmart_liquidacion",
                period, len(guias), resumen.get("total_wcs"), detalle_totales,
            )
    except Exception:
        _get_connection.clear()


def obtener_detalle_totales(liquidacion_id: int) -> dict | None:
    """detalle_totales (JSONB) de una liquidacion puntual. Solo lectura,
    mismo criterio fail-soft que el resto de este modulo."""
    try:
        conn = _get_connection()
        if conn is None:
            return None
        with conn.cursor() as cur:
            cur.execute("SELECT detalle_totales FROM liquidaciones WHERE id = %s", (liquidacion_id,))
            fila = cur.fetchone()
        return fila[0] if fila else None
    except Exception:
        _get_connection.clear()
        return None


def obtener_filtros_historial() -> dict | None:
    """Aerolineas/estaciones/periodos distintos que aparecen en liquidaciones,
    para poblar los selectores de la pantalla de Historial.

    None si la base no esta configurada o no responde (ver docstring del
    modulo) -- la pantalla lo interpreta como "historial no disponible
    ahora", no como un error fatal.
    """
    try:
        conn = _get_connection()
        if conn is None:
            return None
        with conn.cursor() as cur:
            cur.execute("SELECT DISTINCT aerolinea FROM liquidaciones ORDER BY aerolinea")
            aerolineas = [row[0] for row in cur.fetchall()]
            cur.execute("SELECT DISTINCT estacion FROM liquidaciones ORDER BY estacion")
            estaciones = [row[0] for row in cur.fetchall()]
            cur.execute(
                "SELECT DISTINCT periodo, periodo_mes, periodo_anio FROM liquidaciones ORDER BY periodo DESC"
            )
            periodos = cur.fetchall()  # [(date(2026,7,1), 'JUL', '26'), ...]
        return {"aerolineas": aerolineas, "estaciones": estaciones, "periodos": periodos}
    except Exception:
        _get_connection.clear()
        return None


def obtener_historial(
    aerolinea: str | None = None,
    periodo: date | None = None,
    estacion: str | None = None,
) -> pd.DataFrame | None:
    """Liquidaciones filtradas (mas recientes primero). Solo lectura.

    Cada fila es UN reporte generado (ver guardar_reporte_simple /
    guardar_liquidacion_latam): si el mismo periodo/aerolinea/estacion/tipo
    de cargo se genero mas de una vez, aparecen todas aca -- es historial
    real para trazabilidad, no un resumen deduplicado.

    La columna "es_vigente" marca, por combinacion (aerolinea, estacion,
    tipo_cargo, periodo), cual es la fila de generado_en mas reciente
    (True) vs. generaciones anteriores de la misma combinacion (False).
    Quien consuma esto para sumar un total debe filtrar por es_vigente
    para no duplicar plata si algo se regenero -- ver su uso en app.py.

    Un None en cualquier filtro significa "todos".
    """
    try:
        conn = _get_connection()
        if conn is None:
            return None
        # Los ::text/::date son necesarios: sin el cast explicito, Postgres no
        # puede inferir el tipo de un parametro que se usa dos veces (IS NULL
        # y comparacion) y psycopg tira AmbiguousParameter -- probado en el
        # navegador real, no es una precaucion teorica.
        query = """
            SELECT id, carga_id, generado_en, aerolinea, estacion, tipo_cargo,
                   periodo_mes, periodo_anio, periodo, cantidad_filas, monto_total,
                   (ROW_NUMBER() OVER (
                       PARTITION BY aerolinea, estacion, tipo_cargo, periodo
                       ORDER BY generado_en DESC
                   ) = 1) AS es_vigente
            FROM liquidaciones
            WHERE (%(aerolinea)s::text IS NULL OR aerolinea = %(aerolinea)s::text)
              AND (%(periodo)s::date IS NULL OR periodo = %(periodo)s::date)
              AND (%(estacion)s::text IS NULL OR estacion = %(estacion)s::text)
            ORDER BY generado_en DESC
        """
        params = {"aerolinea": aerolinea, "periodo": periodo, "estacion": estacion}
        return pd.read_sql(query, conn, params=params)
    except Exception:
        _get_connection.clear()
        return None


def obtener_movimientos_de_carga(carga_id: int) -> pd.DataFrame | None:
    """Movimientos crudos de una carga, con los mismos nombres de columna
    que trae el archivo original (Codigo, Cod.Vuelo, etc. -- via alias SQL
    a los mismos COL_* de config.py que usa el resto de la app).

    Es deliberado devolverlos con esos nombres: el resultado esta pensado
    para pasarse directo a build_airline_report()/build_latam_detalle()
    (ver su uso en app.py) y asi obtener el mismo detalle que ya calcula
    el Excel real para esa liquidacion, sin reimplementar ningun filtro
    aca -- este modulo solo hace la lectura, la logica de filtrado sigue
    viviendo exclusivamente en report_builder.py/liquidacion_builder.py.

    Solo lectura. None si la base no responde (mismo criterio fail-soft
    que el resto de este modulo).
    """
    try:
        conn = _get_connection()
        if conn is None:
            return None
        query = f"""
            SELECT
                codigo AS "{COL_CODIGO}",
                cod_vuelo AS "{COL_COD_VUELO}",
                cliente AS "{COL_CLIENTE}",
                condicion AS "{COL_CONDICION}",
                tipo AS "{COL_TIPO}",
                tpo_cambio AS "{COL_TPO_CAMBIO}",
                dry_fee AS "{COL_DRY_FEE}",
                aduana AS "{COL_ADUANA}",
                trans_e AS "{COL_TRANS_E}",
                iata AS "{COL_IATA}",
                collect AS "{COL_COLLECT}",
                aerolinea AS "{COL_AEROLINEA}",
                estacion AS "{COL_ESTACION}"
            FROM movimientos_awb
            WHERE carga_id = %(carga_id)s
        """
        movimientos = pd.read_sql(query, conn, params={"carga_id": carga_id})
        # OJO: NO redondear ni castear las columnas de monto a entero aca.
        # Hubo una version anterior que lo hacia asumiendo que siempre son
        # pesos enteros -- FALSO: el archivo real tiene valores con
        # centavos reales incluso en Dry.Fee (ej. 319742.5), y sobre todo
        # en Iata/Collect (usadas por el IVA de LATAM, ej. 851451.42).
        # Redondear ahi corrompia el calculo de build_latam_resumen()
        # (total_periodo terminaba unos pesos distinto del Excel real) --
        # detectado comparando este reconstruido contra el Excel real
        # generado en el momento, antes de dar el fix por bueno. Si en
        # algun momento se quiere mostrar sin ".0" en pantalla, ese
        # formateo tiene que vivir en la capa de presentacion (app.py),
        # nunca aca, para no tocar los valores que alimentan un calculo.
        return movimientos
    except Exception:
        _get_connection.clear()
        return None
