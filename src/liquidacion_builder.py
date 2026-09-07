"""Armado de la liquidacion formal de LATAM.

LATAM no encaja en el flujo simple de report_builder.py (una hoja por tipo
de cargo/estacion): necesita una liquidacion con 3 hojas -- Resumen
Facturacion, Detalle de Facturacion y Compensacion -- con una formula de
sub-facturas e IVA. Por eso vive en un modulo aparte, aunque reutiliza dos
utilidades genericas de report_builder.py (no son especificas del reporte
simple): _nonzero_mask y _fix_codigo_column.

Esta liquidacion esta validada unicamente para LATAM_LIQUIDACION_STATION
(EZE), contra el archivo de referencia real del cliente. No calcula
compensacion (ver nota en config.py junto a LATAM_SUBFACTURA_LA): esa parte
sigue siendo manual, tal como confirmo el cliente.
"""

import re

import pandas as pd
from openpyxl.styles import Font

from src.config import (
    COL_AEROLINEA,
    COL_CLIENTE,
    COL_CODIGO,
    COL_COD_VUELO,
    COL_CONDICION,
    COL_ESTACION,
    COL_TIPO,
    COL_TPO_CAMBIO,
    LATAM_DETALLE_CHARGE_COLUMNS,
    LATAM_IVA_RATE,
    LATAM_LIQUIDACION_STATION,
    LATAM_SUBFACTURA_4M,
    LATAM_SUBFACTURA_LA,
)
from src.report_builder import _fix_codigo_column, _nonzero_mask

# Sufijo de Cod.Vuelo: ej. "LA8130_01JUL26" -> dia=01, mes=JUL, anio=26.
_PERIOD_PATTERN = re.compile(r"_(\d{2})([A-Za-z]{3})(\d{2})$")


def _extract_period(cod_vuelo: object) -> tuple[str, str] | None:
    match = _PERIOD_PATTERN.search(str(cod_vuelo))
    if not match:
        return None
    return match.group(2).upper(), match.group(3)


def detect_period(df: pd.DataFrame) -> tuple[str, str]:
    """Autodetecta el periodo (mes, anio) como el mas frecuente en Cod.Vuelo.

    "archivo original.xlsx" no viene filtrado por periodo: puede traer
    sueltas algunas filas de meses anteriores (vuelos tardios que quedaron
    en el export). Se asume que el periodo real de la liquidacion es el que
    concentra la mayoria de las filas.
    """
    periods = df[COL_COD_VUELO].map(_extract_period).dropna()
    if periods.empty:
        raise ValueError(
            "No se pudo detectar el periodo: ningun Cod.Vuelo tiene el "
            "formato esperado (ej. 'LA8130_01JUL26')."
        )
    return periods.value_counts().idxmax()


def _filter_period(df: pd.DataFrame, period: tuple[str, str]) -> pd.DataFrame:
    mask = df[COL_COD_VUELO].map(_extract_period) == period
    return df[mask]


def build_latam_detalle(df: pd.DataFrame, period: tuple[str, str] | None = None) -> pd.DataFrame:
    """Arma la hoja "Detalle de Facturacion" de LATAM para LATAM_LIQUIDACION_STATION.

    Una fila por guia con los datos identificatorios y las 4 columnas de
    cargo (nombres largos), copiadas tal cual del original sin ningun
    ajuste de tarifa -- LATAM usa el dato bruto, igual que Gol.

    period es (mes, anio) ej. ("JUL", "26"); si es None se autodetecta con
    detect_period().
    """
    latam_df = df[(df[COL_AEROLINEA] == "LATAM") & (df[COL_ESTACION] == LATAM_LIQUIDACION_STATION)]

    if period is None:
        period = detect_period(latam_df)
    latam_df = _filter_period(latam_df, period)

    raw_columns = list(LATAM_DETALLE_CHARGE_COLUMNS.keys())
    mask = _nonzero_mask(latam_df, raw_columns)
    filtered = latam_df.loc[mask].copy()

    id_columns = [COL_CODIGO, COL_COD_VUELO, COL_CLIENTE, COL_CONDICION, COL_TIPO, COL_TPO_CAMBIO]
    detalle = filtered[id_columns + raw_columns].rename(columns=LATAM_DETALLE_CHARGE_COLUMNS)
    detalle = detalle.reset_index(drop=True)
    detalle = _fix_codigo_column(detalle)
    return detalle


def build_latam_resumen(detalle: pd.DataFrame) -> dict:
    """Calcula los totales del Resumen Facturacion a partir del Detalle.

    No aplica ningun ajuste de compensacion: usa los montos en bruto del
    Detalle (ver nota junto a LATAM_SUBFACTURA_LA en config.py).
    """
    sums = {
        col: float(pd.to_numeric(detalle[col], errors="coerce").fillna(0).sum())
        for col in LATAM_DETALLE_CHARGE_COLUMNS.values()
    }

    total_la = sum(sums[c] for c in LATAM_SUBFACTURA_LA)
    neto_gravado = sum(sums[c] for c in LATAM_SUBFACTURA_4M)
    iva = neto_gravado * LATAM_IVA_RATE
    total_4m = neto_gravado + iva
    total_periodo = total_la + total_4m

    return {
        "sums": sums,
        "total_la": total_la,
        "neto_gravado": neto_gravado,
        "iva": iva,
        "total_4m": total_4m,
        "total_periodo": total_periodo,
    }


def _write_resumen_sheet(writer: pd.ExcelWriter, resumen: dict) -> None:
    ws = writer.book.create_sheet("Resumen Facturación")

    bold = Font(bold=True)
    money_format = "#,##0.00"

    def write(row, label, value=None, *, is_bold=False):
        ws.cell(row=row, column=1, value=label)
        if is_bold:
            ws.cell(row=row, column=1).font = bold
        if value is not None:
            cell = ws.cell(row=row, column=3, value=round(value, 2))
            cell.number_format = money_format
            if is_bold:
                cell.font = bold

    sums = resumen["sums"]
    row = 1
    write(row, "LATAM AIRLINES", is_bold=True)
    row += 1
    for col in LATAM_SUBFACTURA_LA:
        write(row, col, sums[col])
        row += 1
    write(row, "TOTAL LA (sin IVA)", resumen["total_la"], is_bold=True)
    row += 2

    write(row, "LAN ARGENTINA", is_bold=True)
    row += 1
    for col in LATAM_SUBFACTURA_4M:
        write(row, col, sums[col])
        row += 1
    write(row, "Total IVA (21%, informativo)", resumen["iva"])
    row += 1
    write(row, "TOTAL 4M (con IVA)", resumen["total_4m"], is_bold=True)
    row += 2

    write(row, "TOTAL PERIODO", resumen["total_periodo"], is_bold=True)

    ws.column_dimensions["A"].width = 65
    ws.column_dimensions["C"].width = 18


def _write_compensacion_sheet(writer: pd.ExcelWriter) -> None:
    ws = writer.book.create_sheet("Compensación")
    ws.cell(row=1, column=1, value="Compensación").font = Font(bold=True, size=13)
    ws.cell(row=3, column=1, value="Cargar manualmente los casos de compensación del período.")
    ws.column_dimensions["A"].width = 65


def write_liquidacion(detalle: pd.DataFrame, resumen: dict, output_path) -> None:
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        _write_resumen_sheet(writer, resumen)
        detalle.to_excel(writer, sheet_name="Detalle de Facturación", index=False)
        _write_compensacion_sheet(writer)
