"""El "mes anterior" de JetSmart, leido de la base en vez de pedir de nuevo
los archivos del mes pasado.

Las reglas del cruce (jetsmart_cruce.py) necesitan saber que guias estaban en
el export y en el Ariel del mes anterior. Este modulo es la parte PURA de ese
trabajo (sin base de datos): arma las filas que se guardan al procesar un mes
y reconstruye, a partir de las filas leidas, los dos DataFrames que espera el
cruce. El acceso a la base esta en db.py (guardar_liquidacion_jetsmart /
obtener_mes_anterior_jetsmart); la tabla, en sql/migrations/002.

Lo que se guarda de cada mes son TODAS las guias de las dos fuentes (las
guias del sistema y la rendicion de Ariel), no solo las que se liquidaron:
las pendientes de declarar y las excluidas son justamente las que el mes
siguiente tiene que reconocer.
"""

from dataclasses import dataclass, field
from datetime import datetime, timedelta, timezone

import pandas as pd

from src.config import (
    COL_JS_CLIENTE,
    COL_JS_CREACION,
    COL_JS_DESTINO,
    COL_JS_EMPRESA,
    COL_JS_ESTADO,
    COL_JS_GUIA,
    COL_JS_KGS,
    COL_JS_ORIGEN,
    COL_JS_PRIORIDAD,
)
from src.jetsmart_cruce import (
    COL_AR_FECHA,
    COL_AR_GUIA,
    COL_AR_INGRESO,
    COL_AR_KG,
    COL_AR_ORIGEN,
)
from src.period_utils import Period, period_label

FUENTE_SISTEMA = "sistema"
FUENTE_ARIEL = "ariel"
ORIGEN_LIQUIDACION = "liquidacion"
ORIGEN_CARGA_INICIAL = "carga_inicial"

# Argentina no tiene horario de verano: UTC-3 fijo, sin depender de tzdata (que no viene en Windows).
_ARGENTINA = timezone(timedelta(hours=-3))


class MesAnteriorNoDisponible(Exception):
    """El mes anterior se habia leido de la base y ya no se pudo leer al procesar."""


@dataclass
class ResumenMesAnterior:
    periodo: Period
    version: int
    creado_en: datetime
    origen_dato: str            # ORIGEN_LIQUIDACION / ORIGEN_CARGA_INICIAL
    n_sistema: int = 0
    n_ariel: int = 0

    @property
    def texto(self) -> str:
        """Lo que se le muestra a la usuaria, ej. "Mes anterior: septiembre 2026,
        versión 2, cargado el 4/11"."""
        return (
            f"Mes anterior: {period_label(self.periodo)}, versión {self.version}, "
            f"cargado el {fecha_corta(self.creado_en)}"
        )


@dataclass
class LecturaMesAnterior:
    """Resultado de leer el mes anterior de la base.
    estado: 'ok' (hay datos), 'sin_datos' (la base responde pero ese mes no
    esta) o 'sin_base' (la base no esta configurada o no responde)."""
    estado: str
    resumen: ResumenMesAnterior | None = None
    filas: list[dict] = field(default_factory=list)   # solo si se pidieron las guias


def fecha_corta(dt: datetime) -> str:
    """4/11 (dia/mes, hora argentina)."""
    local = dt.astimezone(_ARGENTINA) if dt.tzinfo else dt
    return f"{local.day}/{local.month}"


def _texto(valor) -> str | None:
    return None if pd.isna(valor) else str(valor).strip()


def _numero(valor) -> float | None:
    return None if pd.isna(valor) else float(valor)


def _fecha(valor):
    ts = pd.to_datetime(valor, dayfirst=True, errors="coerce")
    return None if pd.isna(ts) else ts.date()


def filas_guias_periodo(export: pd.DataFrame, ariel: pd.DataFrame, destinos: dict | None = None) -> list[dict]:
    """Filas a guardar de un mes: una por guia y por fuente.

    export: guias del sistema en formato interno (COL_JS_*). ariel: guias de
    Ariel ya limpias (COL_AR_*). destinos: {guia: (resultado, motivo)} del
    cruce (CruceResultado.destinos); None en la carga inicial, donde ese mes
    no se cruzo aca y no hay resultado que inventar.

    Una guia aparece una sola vez por fuente: si Ariel trae el mismo numero
    con datos distintos (ej. 97479 de agosto 2026) se guarda la primera fila;
    para el mes siguiente solo importa que el numero estaba.
    """
    destinos = destinos or {}

    def resultado(g):
        return destinos.get(g, (None, None))

    filas = []
    exp = export.assign(_g=pd.to_numeric(export[COL_JS_GUIA], errors="coerce")).dropna(subset=["_g"])
    for _, r in exp.drop_duplicates("_g").iterrows():
        g = int(r["_g"])
        res, motivo = resultado(g)
        filas.append({
            "fuente": FUENTE_SISTEMA, "nro_guia": g, "fecha": _fecha(r[COL_JS_CREACION]),
            "origen": _texto(r[COL_JS_ORIGEN]), "kg": _numero(r[COL_JS_KGS]), "ingreso": _numero(r[COL_JS_PRIORIDAD]),
            "estado": _texto(r[COL_JS_ESTADO]), "resultado": res, "motivo": motivo,
        })
    for _, r in ariel.drop_duplicates(COL_AR_GUIA).iterrows():
        g = int(r[COL_AR_GUIA])
        res, motivo = resultado(g)
        filas.append({
            "fuente": FUENTE_ARIEL, "nro_guia": g, "fecha": _fecha(r[COL_AR_FECHA]),
            "origen": _texto(r[COL_AR_ORIGEN]), "kg": _numero(r[COL_AR_KG]), "ingreso": _numero(r[COL_AR_INGRESO]),
            "estado": None, "resultado": res, "motivo": motivo,
        })
    return filas


def frames_desde_filas(filas: list[dict]) -> tuple[pd.DataFrame, pd.DataFrame]:
    """(export, ariel) del mes anterior, en el formato que espera
    cruzar_jetsmart, reconstruidos desde las filas guardadas. El cruce solo
    usa de ellos el numero de guia (que guias estaban), pero se completan las
    demas columnas guardadas para que sean DataFrames normales."""
    sistema = [f for f in filas if f["fuente"] == FUENTE_SISTEMA]
    ariel = [f for f in filas if f["fuente"] == FUENTE_ARIEL]
    export = pd.DataFrame({
        COL_JS_CREACION: [f["fecha"].strftime("%d/%m/%Y") if f["fecha"] else None for f in sistema],
        COL_JS_GUIA: pd.Series([f["nro_guia"] for f in sistema], dtype="int64"),
        COL_JS_CLIENTE: [None] * len(sistema),
        COL_JS_ORIGEN: [f["origen"] for f in sistema],
        COL_JS_DESTINO: [None] * len(sistema),
        COL_JS_ESTADO: [f["estado"] for f in sistema],
        COL_JS_KGS: [f["kg"] for f in sistema],
        COL_JS_PRIORIDAD: [f["ingreso"] for f in sistema],
        COL_JS_EMPRESA: ["Jetsmart"] * len(sistema),
    })
    ar = pd.DataFrame({
        COL_AR_GUIA: pd.Series([f["nro_guia"] for f in ariel], dtype="int64"),
        COL_AR_FECHA: pd.to_datetime([f["fecha"] for f in ariel]),
        COL_AR_ORIGEN: [f["origen"] for f in ariel],
        COL_AR_KG: [f["kg"] for f in ariel],
        COL_AR_INGRESO: [f["ingreso"] for f in ariel],
    })
    return export, ar
