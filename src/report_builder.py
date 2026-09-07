"""Armado de las hojas del reporte filtrado por aerolinea.

Esta etapa filtra, selecciona y agrupa columnas: no calcula tarifas, splits
USD/ARS ni IVA (eso queda para una segunda etapa una vez confirmada la
tabla de tarifas con el cliente). La unica excepcion es el Delivery Fee de
Avianca, que en las estaciones confirmadas no depende de ninguna tabla de
tarifas sino de un monto fijo por guia (ver
AVIANCA_DELIVERY_FEE_USD_POR_ESTACION en config.py).
"""

import pandas as pd

from src.config import (
    AIRLINE_CONFIGS,
    AVIANCA_DELIVERY_FEE_USD_POR_ESTACION,
    CHARGE_TYPES,
    COL_AEROLINEA,
    COL_CODIGO,
    COL_DRY_FEE,
    COL_ESTACION,
    COL_TPO_CAMBIO,
    EMPTY_SHEET_STYLE_HEADERS_ONLY,
    EMPTY_SHEET_STYLE_PLACEHOLDER,
    EMPTY_STATION_TEXT,
    STATIONS,
)


def _nonzero_mask(df: pd.DataFrame, amount_columns: list[str]) -> pd.Series:
    mask = pd.Series(False, index=df.index)
    for col in amount_columns:
        mask |= pd.to_numeric(df[col], errors="coerce").fillna(0) != 0
    return mask


def _format_codigo(value: object) -> object:
    """Devuelve el AWB como texto plano, sin notacion cientifica ni ".0".

    "Código" mezcla AWBs numericos (Master, hasta 11 digitos: justo el
    umbral en el que Excel pasa el formato General a notacion cientifica,
    ej. 7,3E+10) con AWBs alfanumericos o con ceros a la izquierda (House).
    Escribirla siempre como texto evita ambos problemas sin tocar ni un
    digito del valor real.
    """
    if pd.isna(value):
        return value
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def _fix_codigo_column(df: pd.DataFrame) -> pd.DataFrame:
    if COL_CODIGO in df.columns:
        df = df.copy()
        df[COL_CODIGO] = df[COL_CODIGO].map(_format_codigo)
    return df


def _placeholder_sheet(text: str) -> pd.DataFrame:
    return pd.DataFrame(columns=[text])


def _apply_delivery_fee_override(df: pd.DataFrame, fixed_usd: float | None) -> pd.DataFrame:
    """Reemplaza "Dry.Fee" por un monto fijo en USD (fixed_usd x Tpo.Cambio).

    fixed_usd es el valor confirmado para la estacion de esta hoja (ver
    AVIANCA_DELIVERY_FEE_USD_POR_ESTACION en config.py). Si es None, no hay
    ajuste confirmado para esa estacion todavia: se deja "Dry.Fee" tal cual
    viene del original, sin tocarlo.
    """
    if fixed_usd is None or df.empty or COL_DRY_FEE not in df.columns:
        return df
    df = df.copy()
    tpo_cambio = pd.to_numeric(df[COL_TPO_CAMBIO], errors="coerce")
    df[COL_DRY_FEE] = (fixed_usd * tpo_cambio).round().astype("int64")
    return df


def _empty_sheet(empty_sheet_style: str, report_columns: list[str], placeholder_text: str) -> pd.DataFrame:
    """Arma la hoja para un tipo de cargo/estacion sin ninguna fila.

    "headers_only" arma la hoja vacia con los encabezados reales de columna
    (sin ningun texto), tal como confirmo Cristian para Gol. El default
    "placeholder_text" mantiene el texto explicativo de "sin movimiento".
    """
    if empty_sheet_style == EMPTY_SHEET_STYLE_HEADERS_ONLY:
        return pd.DataFrame(columns=report_columns)
    return _placeholder_sheet(placeholder_text)


def build_airline_report(df: pd.DataFrame, airline_key: str) -> dict[str, pd.DataFrame]:
    """Devuelve {nombre_de_hoja: DataFrame} para la aerolinea pedida.

    airline_key es una clave de AIRLINE_CONFIGS (ej. "avianca", "gol").
    """
    if airline_key not in AIRLINE_CONFIGS:
        raise ValueError(
            f"Aerolinea desconocida: {airline_key!r}. "
            f"Opciones disponibles: {sorted(AIRLINE_CONFIGS)}"
        )

    airline_cfg = AIRLINE_CONFIGS[airline_key]
    airline_df = df[df[COL_AEROLINEA] == airline_cfg["match"]]
    empty_sheet_style = airline_cfg.get("empty_sheet_style", EMPTY_SHEET_STYLE_PLACEHOLDER)

    sheets: dict[str, pd.DataFrame] = {}
    for charge_type_key in airline_cfg["charge_types"]:
        charge_cfg = CHARGE_TYPES[charge_type_key]
        delivery_fee_usd_by_station = None
        if airline_key == "avianca" and charge_type_key == "delivery_fee":
            delivery_fee_usd_by_station = AVIANCA_DELIVERY_FEE_USD_POR_ESTACION
        charge_sheets = _build_charge_type_sheets(airline_df, charge_cfg, empty_sheet_style, delivery_fee_usd_by_station)
        sheets.update(charge_sheets)
    return sheets


def _build_charge_type_sheets(
    airline_df: pd.DataFrame,
    charge_cfg: dict,
    empty_sheet_style: str,
    delivery_fee_usd_by_station: dict | None = None,
) -> dict[str, pd.DataFrame]:
    amount_columns = charge_cfg["amount_columns"]
    report_columns = charge_cfg["report_columns"]

    if charge_cfg["per_station"]:
        result = {}
        for station in STATIONS:
            station_df = airline_df[airline_df[COL_ESTACION] == station]
            mask = _nonzero_mask(station_df, amount_columns)
            filtered = station_df.loc[mask, report_columns].reset_index(drop=True)
            filtered = _fix_codigo_column(filtered)
            if delivery_fee_usd_by_station is not None:
                filtered = _apply_delivery_fee_override(filtered, delivery_fee_usd_by_station.get(station))

            sheet_name = f"{charge_cfg['sheet_prefix']} {station}"
            if filtered.empty:
                placeholder_text = EMPTY_STATION_TEXT.format(station=station)
                filtered = _empty_sheet(empty_sheet_style, report_columns, placeholder_text)
            result[sheet_name] = filtered
        return result

    mask = _nonzero_mask(airline_df, amount_columns)
    filtered = airline_df.loc[mask, report_columns].reset_index(drop=True)
    filtered = _fix_codigo_column(filtered)
    sheet_name = charge_cfg["sheet_name"]
    if filtered.empty:
        placeholder_text = f"Sin movimiento de awbs con cargo {sheet_name}"
        filtered = _empty_sheet(empty_sheet_style, report_columns, placeholder_text)
    return {sheet_name: filtered}


def detect_unconfirmed_activity(sheets: dict[str, pd.DataFrame], unconfirmed_stations: list[str]) -> dict[str, int]:
    """Red de seguridad: avisa si hay movimiento real en una estacion sin validar.

    El reporte simple NO descarta filas de estaciones sin confirmar: las
    incluye igual con el dato bruto (ver AVIANCA_DELIVERY_FEE_USD_POR_ESTACION
    en config.py, que deja "Dry.Fee" sin ajustar para esas estaciones). Esta
    funcion solo detecta si eso paso, para que la interfaz lo muestre como
    advertencia en vez de presentarlo con la misma confianza que EZE/COR.

    Devuelve {estacion: cantidad de filas} solo para las estaciones de
    unconfirmed_stations que tienen al menos una fila en algun sheet.
    """
    activity: dict[str, int] = {}
    for station in unconfirmed_stations:
        n_rows = sum(len(df) for name, df in sheets.items() if name.endswith(f" {station}"))
        if n_rows > 0:
            activity[station] = n_rows
    return activity


def write_report(sheets: dict[str, pd.DataFrame], output_path: str) -> None:
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, sheet_df in sheets.items():
            sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)
