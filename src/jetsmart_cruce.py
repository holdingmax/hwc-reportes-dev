"""Cruce automatico de JetSmart: export de guias de Handyway vs archivo de Ariel.

Reemplaza la comparacion manual de la hoja CONTROL de Anita. Reglas
confirmadas con la operacion (validadas contra las liquidaciones finales de
junio y julio 2026 en tests/test_jetsmart_cruce.py):

- Se cruza por numero de guia exacto: "# Guía" del export contra
  "Nro de guía" de Ariel. Antes de cruzar se descartan las guias anuladas
  de Ariel (texto "ANULADA" en cualquier columna; Ariel reusa el numero de
  una guia anulada para otra valida con otro AWB).
- Ariel es la fuente de la liquidacion: lo que entra se liquida con los Kg,
  ingreso y origen de Ariel.
- Diferencias de Kg o ingreso entre ambos archivos: sin tolerancia (aunque
  sean centavos). La guia entra con los valores de Ariel y se marca.
- Guia solo en el export:
    * si esta en el Ariel del mes anterior -> ya se declaro y liquido ese
      mes: queda afuera, sin excepcion (es el corte de mes normal).
    * si no -> "pendiente de declarar": queda afuera y se marca.
- Guia solo en Ariel: entra si se puede confirmar --
    * esta en el export del mes anterior (era una pendiente de declarar que
      Ariel declara recien ahora), o
    * esta en el export del mes siguiente, si ya se cargo, o
    * si el export del mes siguiente NO se cargo y la fecha de Ariel es el
      ultimo dia calendario del mes: entra igual, con aviso visible de que
      se incluyo por esa regla sin confirmacion cruzada.
  Si esta tambien en el Ariel del mes anterior (declarada dos veces) o no
  hay forma de confirmarla, queda afuera y va a revision manual.

Fuera de alcance: la validacion contra AFIP y las correcciones manuales que
Anita hace sobre su copia del archivo de Ariel (ej. un ingreso corregido a
mano): esas aparecen aca como diferencia marcada, no como error silencioso.
"""

import calendar
import unicodedata
from dataclasses import dataclass, field

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
from src.period_utils import Period

# Columnas del archivo de Ariel (hoja "ARIEL" del archivo de trabajo de Anita).
COL_AR_GUIA = "Nro de guía"
COL_AR_AWB = "AWB"
COL_AR_FECHA = "Fecha de guia"
COL_AR_CLIENTE = "Cliente"
COL_AR_ORIGEN = "Origen"
COL_AR_DESTINO = "Destino"
COL_AR_KG = "Kg"
COL_AR_INGRESO = "Ingreso total flete"
_AR_REQUIRED = [COL_AR_GUIA, COL_AR_FECHA, COL_AR_ORIGEN, COL_AR_KG, COL_AR_INGRESO]

# Columnas de la tabla de Ariel, tal como las trae la hoja "Venta" de su
# rendicion mensual y la hoja "ARIEL" del archivo de trabajo (son las mismas;
# "Prefihjo" asi, con el typo del original).
_AR_COLUMNS = [
    "Prefihjo", COL_AR_AWB, COL_AR_GUIA, COL_AR_FECHA, "Nro de vuelo embarcada", COL_AR_CLIENTE,
    "Destinatario", COL_AR_ORIGEN, COL_AR_DESTINO, "Bultos", COL_AR_KG, "Kg comercial",
    "Tarifa por kg", COL_AR_INGRESO, "Descripción de la mercadería",
]
# Hasta que fila se busca el encabezado de la tabla en la hoja "Venta": el
# archivo de Ariel trae 1 o 2 filas en blanco arriba (varia mes a mes).
_RAW_HEADER_MAX_ROW = 20

_MES_ABREV = ["ENE", "FEB", "MAR", "ABR", "MAY", "JUN", "JUL", "AGO", "SEP", "OCT", "NOV", "DIC"]

# Tipos de excepcion (lo unico que se muestra: lo que matchea limpio no).
EXC_DIFERENCIA = "Diferencia de Kg / ingreso"
EXC_PENDIENTE = "Pendiente de declarar"
EXC_REVISION = "Solo en Ariel, sin confirmar — revisión manual"
EXC_DOBLE = "Declarada también en el Ariel del mes anterior — revisión manual"
EXC_DUPLICADA = "Número de guía repetido en Ariel — revisión manual"

# Motivos de inclusion (para el resumen y el aviso de la regla de ultimo dia).
INC_MATCH = "match"
INC_DIFERENCIA = "diferencia (valores de Ariel)"
INC_PENDIENTE_ANTERIOR = "pendiente del mes anterior, declarada ahora"
INC_EXPORT_SIGUIENTE = "confirmada con el export del mes siguiente"
INC_ULTIMO_DIA = "último día del mes, sin confirmación cruzada"


def _rewind(source) -> None:
    # Los archivos subidos en Streamlit son buffers: hay que volver al
    # principio antes de cada lectura.
    if hasattr(source, "seek"):
        source.seek(0)


def _find_sheet(xls: pd.ExcelFile, name: str) -> str | None:
    return next((s for s in xls.sheet_names if s.strip().upper() == name), None)


def _norm(texto) -> str:
    """Normaliza un encabezado para compararlo: minusculas, sin tildes, sin
    espacios de mas ("Nro de Guía " -> "nro de guia")."""
    texto = unicodedata.normalize("NFKD", str(texto)).encode("ascii", "ignore").decode("ascii")
    return " ".join(texto.lower().split())


_AR_COLUMNS_NORM = {_norm(c): c for c in _AR_COLUMNS}


def _es_fila_encabezado(valores) -> bool:
    normalizados = {_norm(v) for v in valores if pd.notna(v)}
    return _norm(COL_AR_GUIA) in normalizados and bool(normalizados & {"prefihjo", "prefijo"})


def _limpiar_ariel(df: pd.DataFrame, origen: str) -> tuple[pd.DataFrame, int]:
    """Limpieza comun a la hoja ARIEL (archivo de trabajo) y a la hoja Venta
    (rendicion cruda de Ariel). Descarta: filas sin numero de guia numerico
    (pie de totales, filas vacias) y guias anuladas.

    origen: nombre de la hoja, para el mensaje de error.
    """
    faltantes = [c for c in _AR_REQUIRED if c not in df.columns]
    if faltantes:
        raise ValueError(f"La hoja {origen} no tiene las columnas esperadas: faltan " + ", ".join(faltantes))

    anulada = df.astype(str).apply(lambda col: col.str.strip().str.upper().eq("ANULADA")).any(axis=1)
    nro = pd.to_numeric(df[COL_AR_GUIA], errors="coerce")
    n_anuladas = int(anulada.sum())
    df = df[nro.notna() & ~anulada].copy()
    df[COL_AR_GUIA] = df[COL_AR_GUIA].astype("int64")
    df[COL_AR_FECHA] = pd.to_datetime(df[COL_AR_FECHA], errors="coerce")
    df[COL_AR_ORIGEN] = df[COL_AR_ORIGEN].astype(str).str.strip().str.upper()
    return df.reset_index(drop=True), n_anuladas


def has_ariel_sheet(source) -> bool:
    _rewind(source)
    with pd.ExcelFile(source) as xls:
        return _find_sheet(xls, "ARIEL") is not None


def load_ariel(source) -> tuple[pd.DataFrame, int]:
    """(guias validas de Ariel, cantidad de anuladas descartadas).

    Lee la hoja "ARIEL" del archivo de trabajo. Descarta: filas sin numero
    de guia numerico (pie de totales, filas vacias) y guias anuladas.
    """
    _rewind(source)
    with pd.ExcelFile(source) as xls:
        sheet = _find_sheet(xls, "ARIEL")
        if sheet is None:
            raise ValueError("El archivo no tiene la hoja ARIEL (el archivo de Ariel del mes, tal como lo arma Anita).")
        df = pd.read_excel(xls, sheet_name=sheet)
    return _limpiar_ariel(df, "ARIEL")


def _leer_venta(source) -> pd.DataFrame:
    """Tabla de la hoja "Venta" de la rendicion cruda de Ariel, con los
    encabezados de la tabla como columnas. Lanza ValueError si no esta la
    hoja o no aparece el encabezado en las primeras filas."""
    _rewind(source)
    with pd.ExcelFile(source) as xls:
        sheet = _find_sheet(xls, "VENTA")
        if sheet is None:
            raise ValueError(
                "El archivo no tiene la hoja Venta (la rendición que manda Ariel, con la tabla de guías "
                "declaradas). Hojas encontradas: " + ", ".join(xls.sheet_names)
            )
        crudo = pd.read_excel(xls, sheet_name=sheet, header=None)
    # El offset de filas en blanco varia mes a mes: se busca la fila del
    # encabezado ("Prefihjo" + "Nro de guía") en vez de asumir una fija.
    fila = next(
        (i for i in range(min(_RAW_HEADER_MAX_ROW, len(crudo))) if _es_fila_encabezado(crudo.iloc[i])),
        None,
    )
    if fila is None:
        raise ValueError(
            f"La hoja {sheet} no tiene el encabezado esperado (Prefihjo, Nro de guía, ...) en sus primeras "
            f"{_RAW_HEADER_MAX_ROW} filas: no parece la tabla de guías de Ariel."
        )
    encabezado = crudo.iloc[fila]
    columnas = [_AR_COLUMNS_NORM.get(_norm(v), str(v).strip()) if pd.notna(v) else None for v in encabezado]
    df = crudo.iloc[fila + 1:].copy()
    df.columns = columnas
    # Columnas sin encabezado (ej. la tabla arranca en la columna B): afuera.
    return df.loc[:, [c is not None for c in columnas]].reset_index(drop=True)


def has_ariel_raw_file(source) -> bool:
    """True si el archivo parece la rendicion cruda de Ariel (hoja Venta con
    la tabla de guias). No lanza excepciones: es para habilitar la carga."""
    try:
        _leer_venta(source)
        return True
    except Exception:
        return False


def load_ariel_raw(source) -> tuple[pd.DataFrame, int]:
    """Igual que load_ariel, pero leyendo el archivo CRUDO que manda Ariel
    (hoja "Venta", sin importar mayusculas) en vez de la hoja ARIEL del
    archivo de trabajo. Las demas hojas (manifiestos) se ignoran. Devuelve
    el mismo formato: (guias validas, cantidad de anuladas descartadas)."""
    return _limpiar_ariel(_leer_venta(source), "Venta")


def _guias_set(export: pd.DataFrame | None) -> set[int]:
    if export is None:
        return set()
    return set(pd.to_numeric(export[COL_JS_GUIA], errors="coerce").dropna().astype("int64"))


def _last_day(period: Period) -> pd.Timestamp:
    mes, anio = period
    year, month = 2000 + int(anio), _MES_ABREV.index(mes) + 1
    return pd.Timestamp(year, month, calendar.monthrange(year, month)[1])


def _num(value) -> float | None:
    return None if pd.isna(value) else float(value)


@dataclass
class CruceResultado:
    # Guias que entran a la liquidacion, en formato export (listas para
    # build_jetsmart_guias), con los valores de Ariel.
    incluidas: pd.DataFrame
    # Solo excepciones: una fila por guia, con tipo y detalle.
    excepciones: pd.DataFrame
    # Guias incluidas por la regla de ultimo dia (aviso visible).
    incluidas_ultimo_dia: pd.DataFrame
    resumen: dict = field(default_factory=dict)


_EXC_COLUMNS = ["Tipo", "Nro guía", "Fecha Ariel", "Origen", "Kg export", "Kg Ariel",
                "Ingreso export", "Ingreso Ariel", "¿Entra a la liquidación?", "Detalle"]


def cruzar_jetsmart(
    export: pd.DataFrame,
    ariel: pd.DataFrame,
    period: Period,
    export_anterior: pd.DataFrame | None,
    ariel_anterior: pd.DataFrame | None,
    export_siguiente: pd.DataFrame | None = None,
) -> CruceResultado:
    """Cruza el export del mes contra el Ariel del mes (ver reglas arriba).

    export / export_anterior / export_siguiente: como los devuelve
    load_jetsmart_export. ariel / ariel_anterior: como los devuelve
    load_ariel. export_siguiente es opcional.
    """
    exp = export.copy()
    exp["_g"] = pd.to_numeric(exp[COL_JS_GUIA], errors="coerce").astype("int64")
    exp = exp.drop_duplicates("_g")
    exp_by_g = exp.set_index("_g")

    dup_mask = ariel[COL_AR_GUIA].duplicated(keep=False)
    duplicadas = ariel[dup_mask]
    ar = ariel[~dup_mask]
    ar_by_g = ar.set_index(COL_AR_GUIA)

    prev_ariel = set(ariel_anterior[COL_AR_GUIA]) if ariel_anterior is not None else set()
    prev_export = _guias_set(export_anterior)
    next_export = _guias_set(export_siguiente)
    hay_siguiente = export_siguiente is not None
    ultimo_dia = _last_day(period)

    incluidas, excepciones, ultimo_dia_rows = [], [], []
    motivos: dict[str, int] = {}
    ya_declaradas = 0

    def excepcion(tipo, g, a=None, e=None, entra="No", detalle=""):
        excepciones.append({
            "Tipo": tipo,
            "Nro guía": int(g),
            "Fecha Ariel": a[COL_AR_FECHA].strftime("%d/%m/%Y") if a is not None and pd.notna(a[COL_AR_FECHA]) else None,
            "Origen": (a[COL_AR_ORIGEN] if a is not None else str(e[COL_JS_ORIGEN]).strip().upper()),
            "Kg export": _num(e[COL_JS_KGS]) if e is not None else None,
            "Kg Ariel": _num(a[COL_AR_KG]) if a is not None else None,
            "Ingreso export": _num(e[COL_JS_PRIORIDAD]) if e is not None else None,
            "Ingreso Ariel": _num(a[COL_AR_INGRESO]) if a is not None else None,
            "¿Entra a la liquidación?": entra,
            "Detalle": detalle,
        })

    def incluir(g, a, e, motivo):
        motivos[motivo] = motivos.get(motivo, 0) + 1
        incluidas.append({
            COL_JS_CREACION: (e[COL_JS_CREACION] if e is not None
                              else a[COL_AR_FECHA].strftime("%d/%m/%Y") if pd.notna(a[COL_AR_FECHA]) else None),
            COL_JS_GUIA: int(g),
            COL_JS_CLIENTE: a.get(COL_AR_CLIENTE, e[COL_JS_CLIENTE] if e is not None else None),
            COL_JS_ORIGEN: a[COL_AR_ORIGEN],
            COL_JS_DESTINO: a.get(COL_AR_DESTINO, e[COL_JS_DESTINO] if e is not None else None),
            COL_JS_ESTADO: e[COL_JS_ESTADO] if e is not None else "Solo en Ariel",
            COL_JS_KGS: a[COL_AR_KG],
            COL_JS_PRIORIDAD: a[COL_AR_INGRESO],
            COL_JS_EMPRESA: "Jetsmart",
        })

    for g in sorted(set(exp_by_g.index) | set(ar_by_g.index)):
        e = exp_by_g.loc[g] if g in exp_by_g.index else None
        a = ar_by_g.loc[g] if g in ar_by_g.index else None
        if a is not None:
            a = a.copy()
            a[COL_AR_GUIA] = g

        if e is not None and a is not None:
            same_kg = _num(e[COL_JS_KGS]) == _num(a[COL_AR_KG])
            same_ing = _num(e[COL_JS_PRIORIDAD]) == _num(a[COL_AR_INGRESO])
            if same_kg and same_ing:
                incluir(g, a, e, INC_MATCH)
            else:
                que = " y ".join(x for x, ok in (("Kg", same_kg), ("ingreso", same_ing)) if not ok)
                incluir(g, a, e, INC_DIFERENCIA)
                excepcion(EXC_DIFERENCIA, g, a, e, entra="Sí, con valores de Ariel", detalle=f"Difiere {que}.")
        elif e is not None:
            if g in duplicadas[COL_AR_GUIA].values:
                excepcion(EXC_DUPLICADA, g, None, e, detalle="Ariel trae este número más de una vez (sin contar anuladas).")
            elif g in prev_ariel:
                ya_declaradas += 1
            else:
                excepcion(EXC_PENDIENTE, g, None, e, detalle="No está en el Ariel de este mes ni del anterior: probablemente se declare el mes que viene.")
        else:
            if g in prev_ariel:
                excepcion(EXC_DOBLE, g, a, None, detalle="Ya se declaró el mes anterior: no se incluye para no liquidarla dos veces.")
            elif g in prev_export:
                incluir(g, a, None, INC_PENDIENTE_ANTERIOR)
            elif g in next_export:
                incluir(g, a, None, INC_EXPORT_SIGUIENTE)
            elif not hay_siguiente and pd.notna(a[COL_AR_FECHA]) and a[COL_AR_FECHA].normalize() == ultimo_dia:
                incluir(g, a, None, INC_ULTIMO_DIA)
                ultimo_dia_rows.append({
                    "Nro guía": int(g), "Fecha Ariel": a[COL_AR_FECHA].strftime("%d/%m/%Y"),
                    "Origen": a[COL_AR_ORIGEN], "Kg Ariel": _num(a[COL_AR_KG]), "Ingreso Ariel": _num(a[COL_AR_INGRESO]),
                })
            else:
                donde = "ni en el export del mes siguiente" if hay_siguiente else "y no es del último día del mes"
                excepcion(EXC_REVISION, g, a, None,
                          detalle=f"No está en el export de este mes, ni en el del mes anterior, {donde}.")

    for g in sorted(set(duplicadas[COL_AR_GUIA]) - set(exp_by_g.index)):
        excepcion(EXC_DUPLICADA, g, duplicadas[duplicadas[COL_AR_GUIA] == g].iloc[0], None,
                  detalle="Ariel trae este número más de una vez (sin contar anuladas).")

    exc_df = pd.DataFrame(excepciones, columns=_EXC_COLUMNS)
    no_incluidas = exc_df[exc_df["¿Entra a la liquidación?"] == "No"]
    resumen = {
        "guias_export": len(exp),
        "guias_ariel": len(ariel),
        "incluidas": len(incluidas),
        "motivos_inclusion": motivos,
        "ya_declaradas_mes_anterior": ya_declaradas,
        "excepciones": len(exc_df),
        "excepciones_por_tipo": exc_df["Tipo"].value_counts().to_dict(),
        "ingreso_fuera_por_revision": float(pd.to_numeric(no_incluidas["Ingreso Ariel"], errors="coerce").fillna(0).sum()),
        "export_siguiente_cargado": hay_siguiente,
        "ultimo_dia": ultimo_dia.strftime("%d/%m/%Y"),
    }
    return CruceResultado(
        incluidas=pd.DataFrame(incluidas, columns=[
            COL_JS_CREACION, COL_JS_GUIA, COL_JS_CLIENTE, COL_JS_ORIGEN, COL_JS_DESTINO,
            COL_JS_ESTADO, COL_JS_KGS, COL_JS_PRIORIDAD, COL_JS_EMPRESA,
        ]),
        excepciones=exc_df,
        incluidas_ultimo_dia=pd.DataFrame(ultimo_dia_rows, columns=["Nro guía", "Fecha Ariel", "Origen", "Kg Ariel", "Ingreso Ariel"]),
        resumen=resumen,
    )
