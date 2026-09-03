"""Carga y limpieza de "archivo original.xlsx"."""

import pandas as pd

from src.config import COL_AEROLINEA


def load_original(path: str) -> pd.DataFrame:
    """Lee el export del sistema y lo deja listo para filtrar.

    - Descarta la columna "Unnamed: 0" (siempre vacia en el export).
    - Descarta la fila de totales del pie de la planilla (no tiene
      Aerolinea cargada; las columnas de monto vienen ahi como texto
      formateado, ej. "$ 552.154.440,50").
    - No fuerza tipos en "Código": mezcla AWBs numéricos (Master) con
      alfanuméricos / con ceros a la izquierda (House), y conservarlos
      tal cual evita perder esa información.
    """
    df = pd.read_excel(path, sheet_name=0)
    df = df.drop(columns=[c for c in df.columns if str(c).startswith("Unnamed:")])
    df = df[df[COL_AEROLINEA].notna()].reset_index(drop=True)
    return df
