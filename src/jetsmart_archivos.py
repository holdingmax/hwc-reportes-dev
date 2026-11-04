"""Reconocimiento de los archivos que se suben para liquidar JetSmart.

Anita no tiene que saber que casillero va con cada archivo: sube todo junto,
en cualquier orden, y este modulo reconoce QUE es cada uno mirando sus hojas
y columnas (nunca el nombre del archivo, que cambia de un mes a otro).

Tipos que reconoce:
- guias del sistema (Handyway, Consulta de guias, todo el mes)
- listado de facturacion (Handyway; solo trae guias ya facturadas)
- rendicion de Ariel (hoja "Venta", cualquier mayuscula)
- archivo de trabajo de Anita (formato anterior: export en la primera hoja
  + hoja ARIEL). Se acepta para no romper su flujo de siempre.
- libro IVA Ventas de AFIP (hoja "Comprobantes")

Reglas de este modulo (todas por motivo de negocio, ver comentarios):
- NUNCA lanza excepciones por un archivo raro: devuelve un resultado con el
  problema explicado en lenguaje de negocio. El usuario nunca ve un
  ValueError ni un traceback por subir el archivo equivocado.
- Solo lee: no guarda nada ni toca la base.
"""

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
from src.jetsmart_builder import detect_jetsmart_period, load_jetsmart_export
from src.jetsmart_cruce import _MES_ABREV, ArielLeido, _find_sheet, _norm, _rewind, leer_ariel, load_ariel
from src.period_utils import Period, period_label

TIPO_GUIAS = "guias_sistema"
TIPO_LISTADO = "listado_facturacion"
TIPO_ARIEL = "ariel"
TIPO_TRABAJO = "archivo_trabajo"
TIPO_IVA = "libro_iva"
TIPO_DESCONOCIDO = "desconocido"

# Nombre de cada tipo con las palabras de Anita (es lo unico que se muestra).
ETIQUETAS = {
    TIPO_GUIAS: "Guías del sistema",
    TIPO_LISTADO: "Listado de facturación",
    TIPO_ARIEL: "Rendición de Ariel",
    TIPO_TRABAJO: "Archivo de trabajo (con la hoja de Ariel)",
    TIPO_IVA: "Libro IVA Ventas de AFIP",
}

# Columnas del export "Guias Aerea" del sistema que se usan -> nombre interno
# (las del listado de facturacion, que es el formato con el que se armo todo
# el cruce). Anita (4/11/2026) usa de ese archivo SOLO: numero de guia, kg,
# prioridad (ingreso), origen y estado; el resto lo descarta. Se agrega la
# fecha de creacion, que el sistema usa para filtrar el mes y la liquidacion
# guarda en cada guia. Cliente y destino NO se leen: la liquidacion los toma
# de Ariel (la fuente de lo que se liquida).
# El numero de guia es la columna "#": "Nro." trae origen + numero
# ("CRD-102178") y su prefijo NO coincide con el origen de la guia, asi que
# no sirve para cruzar contra Ariel.
_GUIAS_A_INTERNO = {
    "#": COL_JS_GUIA,
    "kgs.": COL_JS_KGS,
    "$ prioridad": COL_JS_PRIORIDAD,
    "aero.ori.": COL_JS_ORIGEN,
    "estado": COL_JS_ESTADO,
    "cre.": COL_JS_CREACION,
}
# Con estas seis columnas se reconoce el archivo (el listado de facturacion no
# tiene "#" ni "Kgs." con punto, asi que no se confunden).
_GUIAS_FIRMA = set(_GUIAS_A_INTERNO)
# Si el archivo trae la columna Aerolinea se filtra por JetSmart; si no, se
# asume que ya viene solo JetSmart (asi lo baja Anita).
_GUIAS_COL_AEROLINEA = "aerolinea"
_LISTADO_FIRMA = {_norm(c) for c in (COL_JS_CREACION, COL_JS_GUIA, COL_JS_KGS, COL_JS_PRIORIDAD, COL_JS_EMPRESA)}
_IVA_FIRMA = {"punto de venta", "numero desde"}

_ACEPTADOS = (
    "las guías del sistema, el listado de facturación, la rendición de Ariel o el Libro IVA de AFIP"
)


@dataclass
class ArchivoDetectado:
    nombre: str
    tipo: str
    cantidad: int | None = None          # guias (o comprobantes, en el Libro IVA)
    periodo: Period | None = None        # mes al que pertenece, por sus fechas
    notas: list[str] = field(default_factory=list)   # detalle para mostrar (anuladas, otras aerolineas...)
    problema: str | None = None          # solo si tipo == desconocido: que paso, en lenguaje de negocio

    @property
    def reconocido(self) -> bool:
        return self.tipo != TIPO_DESCONOCIDO

    @property
    def etiqueta(self) -> str:
        return ETIQUETAS.get(self.tipo, "Archivo no reconocido")

    @property
    def resumen(self) -> str:
        """Texto "Detecté: ..." para mostrarle al usuario."""
        if not self.reconocido:
            return self.problema or "No reconozco este archivo."
        unidad = "comprobantes" if self.tipo == TIPO_IVA else "guías"
        partes = [self.etiqueta]
        if self.periodo:
            partes[0] += f" de {period_label(self.periodo)}"
        if self.cantidad is not None:
            partes.append(f"{_miles(self.cantidad)} {unidad}")
        txt = "Detecté: " + ", ".join(partes)
        return txt + (" · " + " · ".join(self.notas) if self.notas else "")


def _miles(n: int) -> str:
    return f"{n:,}".replace(",", ".")


def _periodo_por_fechas(fechas: pd.Series) -> Period | None:
    """Mes dominante de una serie de fechas (texto dia/mes/anio o fecha real).
    None si ninguna es valida."""
    try:
        df = pd.DataFrame({COL_JS_CREACION: fechas})
        return detect_jetsmart_period(df)[0]
    except ValueError:
        return None


def _desconocido(nombre: str, problema: str) -> ArchivoDetectado:
    return ArchivoDetectado(nombre=nombre, tipo=TIPO_DESCONOCIDO, problema=problema)


def _columnas_norm(df: pd.DataFrame) -> set[str]:
    return {_norm(c) for c in df.columns}


def _columnas_export() -> list[str]:
    return [
        COL_JS_CREACION, COL_JS_GUIA, COL_JS_CLIENTE, COL_JS_ORIGEN, COL_JS_DESTINO,
        COL_JS_ESTADO, COL_JS_KGS, COL_JS_PRIORIDAD, COL_JS_EMPRESA,
    ]


def cargar_guias_sistema(source) -> pd.DataFrame:
    """Guias del sistema ("Guias Aerea") en el formato interno del export
    (mismas columnas que load_jetsmart_export), listas para el cruce. Solo se
    leen las columnas que Anita usa (ver _GUIAS_A_INTERNO); cliente y destino
    quedan vacios porque la liquidacion los toma de Ariel.

    Si el archivo trae la columna Aerolinea, filtra por JetSmart sin
    distinguir mayusculas (igual que el listado): si algun dia el sistema
    exporta varias aerolineas juntas, solo entran las de JetSmart.
    """
    _rewind(source)
    df = pd.read_excel(source, sheet_name=0)
    por_norm = {_norm(c): c for c in df.columns}
    faltan = [k for k in _GUIAS_A_INTERNO if k not in por_norm]
    if faltan:
        raise ValueError("faltan columnas de las guías del sistema: " + ", ".join(faltan))
    out = pd.DataFrame({interno: df[por_norm[k]] for k, interno in _GUIAS_A_INTERNO.items()})
    out[COL_JS_CLIENTE] = None
    out[COL_JS_DESTINO] = None
    out[COL_JS_EMPRESA] = "Jetsmart"
    if _GUIAS_COL_AEROLINEA in por_norm:
        out = out[df[por_norm[_GUIAS_COL_AEROLINEA]].astype(str).str.strip().str.upper().eq("JETSMART").to_numpy()]
    out = out[out[COL_JS_GUIA].notna()]
    # Mismo orden de columnas que el listado: las dos fuentes son intercambiables para el cruce.
    return out[_columnas_export()].reset_index(drop=True)


def cargar_export(source, tipo: str) -> pd.DataFrame:
    """Las guias del mes (o de otro mes) en formato interno, segun el tipo
    detectado: guias del sistema, listado o archivo de trabajo."""
    if tipo == TIPO_GUIAS:
        return cargar_guias_sistema(source)
    return load_jetsmart_export(source)


def cargar_ariel(source, tipo: str) -> ArielLeido:
    """Guias de Ariel (con lo descartado): de la rendicion cruda (hoja Venta)
    o de la hoja ARIEL del archivo de trabajo."""
    return leer_ariel(source, cruda=(tipo == TIPO_ARIEL))


def detectar_archivo(source, nombre: str) -> ArchivoDetectado:
    """Reconoce que es el archivo por sus hojas y columnas. No lanza
    excepciones: si no se puede abrir o no es ninguno de los esperados,
    devuelve tipo desconocido con el motivo explicado."""
    try:
        _rewind(source)
        xls = pd.ExcelFile(source)
    except Exception:
        return _desconocido(
            nombre, "No pude abrir este archivo. Tiene que ser un Excel (.xlsx), tal como sale del sistema o como lo manda Ariel."
        )
    with xls:
        hojas = list(xls.sheet_names)
        try:
            return _detectar(source, nombre, xls, hojas)
        except Exception:
            # Red de seguridad: un formato inesperado nunca llega al usuario como traceback.
            return _desconocido(
                nombre,
                f"No pude leer este archivo. Esperaba {_ACEPTADOS}. Hojas que tiene: " + ", ".join(h.strip() for h in hojas) + ".",
            )


def _detectar(source, nombre: str, xls: pd.ExcelFile, hojas: list[str]) -> ArchivoDetectado:
    # 1) Rendicion de Ariel: hoja Venta con el encabezado de su tabla. Va
    # primero porque su hoja no tiene encabezado en la fila 1 y las demas
    # pruebas (que leen la primera fila) no la reconocerian.
    try:
        leido = leer_ariel(source, cruda=True)
    except ValueError:
        leido = None
    if leido is not None:
        ariel, anuladas = leido.guias, leido.anuladas
        # Las anuladas (Ariel reusa su numero para otra guia) no cuentan: se avisa para que el total cierre.
        notas = []
        if anuladas:
            notas.append("1 anulada, no se cuenta" if anuladas == 1 else f"{anuladas} anuladas, no se cuentan")
        # Filas identicas repetidas: son la misma guia cargada dos veces; se cuentan una vez y se avisa.
        if leido.repetidas_identicas:
            n = len(leido.repetidas_identicas)
            notas.append(
                f"{n} {'fila repetida idéntica, se cuenta' if n == 1 else 'filas repetidas idénticas, se cuentan'} una sola vez"
            )
        return ArchivoDetectado(
            nombre, TIPO_ARIEL, cantidad=int(ariel["Nro de guía"].nunique()),
            periodo=_periodo_por_fechas(ariel["Fecha de guia"]), notas=notas,
        )

    # 2) Libro IVA de AFIP: hoja Comprobantes.
    hoja_iva = _find_sheet(xls, "COMPROBANTES")
    if hoja_iva is not None:
        iva = pd.read_excel(xls, sheet_name=hoja_iva)
        if _IVA_FIRMA <= _columnas_norm(iva):
            fecha = next((c for c in iva.columns if _norm(c).startswith("fecha")), None)
            return ArchivoDetectado(
                nombre, TIPO_IVA, cantidad=int(len(iva.dropna(how="all"))),
                periodo=_periodo_por_fechas(iva[fecha]) if fecha else None,
            )

    # 3) Archivos con la tabla en la primera hoja.
    primera = pd.read_excel(xls, sheet_name=0)
    cols = _columnas_norm(primera)

    if _GUIAS_FIRMA <= cols:
        guias = cargar_guias_sistema(source)
        notas = []
        if len(guias) < primera[primera.columns[0]].notna().sum():
            notas.append("las guías de otras aerolíneas no se cuentan")
        return ArchivoDetectado(
            nombre, TIPO_GUIAS, cantidad=int(len(guias)),
            periodo=_periodo_por_fechas(guias[COL_JS_CREACION]), notas=notas,
        )

    if _LISTADO_FIRMA <= cols:
        export = load_jetsmart_export(source)
        periodo = _periodo_por_fechas(export[COL_JS_CREACION])
        if _find_sheet(xls, "ARIEL") is not None:
            # Archivo de trabajo de Anita: la primera hoja es el export y trae ademas la hoja de Ariel.
            ariel, _ = load_ariel(source)
            return ArchivoDetectado(
                nombre, TIPO_TRABAJO, cantidad=int(ariel["Nro de guía"].nunique()),
                periodo=_periodo_por_fechas(ariel["Fecha de guia"]) or periodo,
                notas=["trae también las guías de la hoja de Ariel"],
            )
        return ArchivoDetectado(
            nombre, TIPO_LISTADO, cantidad=int(len(export)), periodo=periodo,
            notas=["solo trae guías ya facturadas"],
        )

    # 4) Un archivo que SI es de la herramienta pero no es de entrada.
    if _find_sheet(xls, "LIQUIDACION") is not None and _find_sheet(xls, "GUIAS") is not None:
        return _desconocido(
            nombre,
            "Este archivo es una liquidación ya armada (tiene las hojas GUIAS y LIQUIDACION), no un archivo de origen. "
            f"Subí {_ACEPTADOS}.",
        )

    return _desconocido(
        nombre,
        f"No reconozco este archivo. Esperaba {_ACEPTADOS}. Este tiene las hojas: "
        + ", ".join(h.strip() for h in hojas) + ".",
    )


# ---------------------------------------------------------------------------
# Armado del conjunto: que archivo cumple cada rol en la liquidacion del mes
# ---------------------------------------------------------------------------

# Preferencia cuando hay mas de un archivo que puede cumplir el mismo rol.
# Guias del sistema > archivo de trabajo > listado: el listado solo trae
# guias ya facturadas, asi que como base del cruce genera falsos "solo en
# Ariel" (77 guias en octubre 2026) y solo se usa si no hay otra cosa.
_PRIO_EXPORT = {TIPO_GUIAS: 3, TIPO_TRABAJO: 2, TIPO_LISTADO: 1}
# Rendicion cruda de Ariel > hoja ARIEL del archivo de trabajo (que ademas
# puede traer correcciones manuales de Anita).
_PRIO_ARIEL = {TIPO_ARIEL: 2, TIPO_TRABAJO: 1}


def _orden(p: Period) -> int:
    return int(p[1]) * 12 + _MES_ABREV.index(p[0])


def _desde_orden(n: int) -> Period:
    return (_MES_ABREV[n % 12], f"{n // 12:02d}")


@dataclass
class ConjuntoJetSmart:
    periodo: Period | None = None
    periodo_anterior: Period | None = None
    export_mes: ArchivoDetectado | None = None
    ariel_mes: ArchivoDetectado | None = None
    export_anterior: ArchivoDetectado | None = None
    ariel_anterior: ArchivoDetectado | None = None
    export_siguiente: ArchivoDetectado | None = None   # opcional: confirma las guias del ultimo dia
    listado: ArchivoDetectado | None = None            # reservado: control contra AFIP (aun no se usa)
    libro_iva: ArchivoDetectado | None = None          # reservado: control contra AFIP (aun no se usa)
    avisos: list[str] = field(default_factory=list)    # no bloquean
    errores: list[str] = field(default_factory=list)   # archivos que no se pudieron usar
    faltantes: list[str] = field(default_factory=list)  # lo que falta para poder procesar

    @property
    def listo(self) -> bool:
        return not self.faltantes


def _mejor(candidatos: list[tuple[int, ArchivoDetectado]], prio: dict[str, int]):
    """El candidato de mayor preferencia (a igual preferencia, el ultimo
    subido) y los que quedaron del mismo tipo que el elegido (duplicados)."""
    if not candidatos:
        return None, []
    elegido = max(candidatos, key=lambda c: (prio[c[1].tipo], c[0]))[1]
    repetidos = [a for _, a in candidatos if a is not elegido and a.tipo == elegido.tipo]
    return elegido, repetidos


def armar_conjunto(archivos: list[ArchivoDetectado]) -> ConjuntoJetSmart:
    """Asigna cada archivo reconocido a su rol y dice que falta, todo en
    lenguaje de negocio.

    Mes a liquidar: el del Ariel mas reciente (Ariel es la fuente de la
    liquidacion: lo que declara es lo que se liquida; las guias del sistema
    se crean por mes de creacion y pueden incluir el mes siguiente). Sin
    ningun Ariel, el de las guias mas recientes.
    El mes anterior se pide aca hasta que se lea de la base de datos.
    """
    c = ConjuntoJetSmart()
    reconocidos = [(i, a) for i, a in enumerate(archivos) if a.reconocido]
    for a in archivos:
        if not a.reconocido:
            c.errores.append(f"{a.nombre}: {a.problema}")

    exports = [(i, a) for i, a in reconocidos if a.tipo in _PRIO_EXPORT]
    ariels = [(i, a) for i, a in reconocidos if a.tipo in _PRIO_ARIEL]
    c.listado = next((a for _, a in reversed(reconocidos) if a.tipo == TIPO_LISTADO), None)
    c.libro_iva = next((a for _, a in reversed(reconocidos) if a.tipo == TIPO_IVA), None)
    if c.libro_iva is not None:
        c.avisos.append(
            f"Reconocí el Libro IVA de AFIP ({c.libro_iva.nombre}). Todavía no se usa en la liquidación: "
            "el control contra AFIP se suma más adelante."
        )

    # Archivos reconocidos pero sin fechas legibles: no se puede saber de que mes son.
    sin_mes = [a for _, a in exports + ariels if a.periodo is None]
    for a in dict.fromkeys(sin_mes):
        c.avisos.append(f"No pude saber de qué mes es {a.nombre} (no encontré fechas): no lo uso.")
    exports = [(i, a) for i, a in exports if a.periodo is not None]
    ariels = [(i, a) for i, a in ariels if a.periodo is not None]

    meses = [a.periodo for _, a in ariels] or [a.periodo for _, a in exports]
    if not meses:
        c.faltantes.append("Subir las guías del mes (las que bajás del sistema) y la rendición de Ariel.")
        return c
    c.periodo = max(meses, key=_orden)
    n = _orden(c.periodo)
    c.periodo_anterior, siguiente = _desde_orden(n - 1), _desde_orden(n + 1)
    mes, ant = period_label(c.periodo), period_label(c.periodo_anterior)

    def rol(pool, periodo, prio):
        elegido, repetidos = _mejor([(i, a) for i, a in pool if a.periodo == periodo], prio)
        for r in repetidos:
            # Dos archivos del mismo tipo y mes: no se pregunta, se usa el ultimo, pero se dice.
            c.avisos.append(
                f"Subiste dos archivos de {r.etiqueta.lower()} de {period_label(periodo)}: "
                f"uso {elegido.nombre} y dejo {r.nombre}."
            )
        return elegido

    c.export_mes = rol(exports, c.periodo, _PRIO_EXPORT)
    c.ariel_mes = rol(ariels, c.periodo, _PRIO_ARIEL)
    c.export_anterior = rol(exports, c.periodo_anterior, _PRIO_EXPORT)
    c.ariel_anterior = rol(ariels, c.periodo_anterior, _PRIO_ARIEL)
    c.export_siguiente = rol(exports, siguiente, _PRIO_EXPORT)

    if c.export_mes is not None and c.export_mes.tipo == TIPO_LISTADO:
        c.avisos.append(
            "Estoy usando el listado de facturación como guías del mes. Solo trae las guías ya facturadas: "
            "puede marcar como 'solo en Ariel' guías que sí existen. Subí las guías del sistema para un cruce completo."
        )
    if c.export_siguiente is not None:
        c.avisos.append(f"Usé las guías de {period_label(siguiente)} para confirmar las guías del último día de {mes}.")

    # Archivos de otros meses: no se ignoran en silencio.
    usados = {id(a) for a in (c.export_mes, c.ariel_mes, c.export_anterior, c.ariel_anterior, c.export_siguiente) if a}
    for _, a in exports + ariels:
        if id(a) not in usados and a.periodo not in (c.periodo, c.periodo_anterior, siguiente):
            c.avisos.append(f"{a.nombre} es de {period_label(a.periodo)}: no lo uso para liquidar {mes}.")

    if c.export_mes is None:
        c.faltantes.append(
            f"Falta el archivo de guías de {mes} (el que bajás del sistema: Consulta de guías, todo el mes)."
        )
    if c.ariel_mes is None:
        c.faltantes.append(f"Falta la rendición de Ariel de {mes} (el Excel que te manda Ariel).")
    if c.export_anterior is None:
        c.faltantes.append(f"Falta el mes anterior: subí las guías de {ant} (el archivo del sistema de ese mes).")
    if c.ariel_anterior is None:
        c.faltantes.append(f"Falta el mes anterior: subí la rendición de Ariel de {ant}.")
    return c
