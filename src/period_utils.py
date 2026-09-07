"""Utilidades de periodo (mes/anio), compartidas por report_builder.py y
liquidacion_builder.py.

"archivo original.xlsx" no viene filtrado por periodo: puede traer sueltas
algunas filas de meses anteriores (vuelos tardios que quedaron en el
export a destiempo). Estas funciones detectan el periodo dominante a
partir del sufijo de Cod.Vuelo (ej. "LA8130_01JUL26" -> mes=JUL, anio=26)
y separan el resto.
"""

import re

import pandas as pd

from src.config import COL_COD_VUELO

# Sufijo de Cod.Vuelo: ej. "LA8130_01JUL26" -> dia=01, mes=JUL, anio=26.
_PERIOD_PATTERN = re.compile(r"_(\d{2})([A-Za-z]{3})(\d{2})$")

_MONTH_NAMES_ES = {
    "ENE": "enero", "FEB": "febrero", "MAR": "marzo", "ABR": "abril",
    "MAY": "mayo", "JUN": "junio", "JUL": "julio", "AGO": "agosto",
    "SEP": "septiembre", "OCT": "octubre", "NOV": "noviembre", "DIC": "diciembre",
    # variantes en ingles, por si el sistema llegara a exportar con esa abreviatura
    "JAN": "enero", "APR": "abril", "AUG": "agosto", "DEC": "diciembre",
}

Period = tuple[str, str]


def extract_period(cod_vuelo: object) -> Period | None:
    match = _PERIOD_PATTERN.search(str(cod_vuelo))
    if not match:
        return None
    return match.group(2).upper(), match.group(3)


def detect_period(df: pd.DataFrame) -> Period:
    """Autodetecta el periodo (mes, anio) como el mas frecuente en Cod.Vuelo.

    Se asume que el periodo real del archivo es el que concentra la
    mayoria de las filas.
    """
    periods = df[COL_COD_VUELO].map(extract_period).dropna()
    if periods.empty:
        raise ValueError(
            "No se pudo detectar el periodo: ningun Cod.Vuelo tiene el "
            "formato esperado (ej. 'LA8130_01JUL26')."
        )
    return periods.value_counts().idxmax()


def filter_by_period(df: pd.DataFrame, period: Period) -> pd.DataFrame:
    """Filas que SI pertenecen al periodo."""
    mask = df[COL_COD_VUELO].map(extract_period) == period
    return df[mask]


def excluded_by_period(df: pd.DataFrame, period: Period) -> pd.DataFrame:
    """Filas que NO pertenecen al periodo (para poder avisar cuantas quedaron afuera)."""
    mask = df[COL_COD_VUELO].map(extract_period) == period
    return df[~mask]


def period_label(period: Period) -> str:
    """Nombre legible del periodo, ej. ("JUL", "26") -> "julio 2026"."""
    month, year = period
    month_name = _MONTH_NAMES_ES.get(month, month.title())
    return f"{month_name} 20{year}"


def period_slug(period: Period) -> str:
    """Slug para nombre de archivo, ej. ("JUL", "26") -> "julio2026"."""
    month, year = period
    month_name = _MONTH_NAMES_ES.get(month, month.lower())
    return f"{month_name}20{year}"
