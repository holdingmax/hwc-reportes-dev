"""CLI para generar el reporte filtrado de una aerolinea.

Uso:
    python -m src.main --airline avianca
    python -m src.main --airline gol --input "data/archivo original.xlsx" --output "output/gol.xlsx"
    python -m src.main --airline latam
"""

import argparse

from src.config import AIRLINE_CONFIGS, FLOW_LIQUIDACION
from src.liquidacion_builder import build_latam_detalle, build_latam_resumen, write_liquidacion
from src.loader import load_original
from src.report_builder import build_airline_report, write_report


def main() -> None:
    parser = argparse.ArgumentParser(description="Genera el reporte filtrado de una aerolinea a partir del export del sistema.")
    parser.add_argument("--airline", required=True, choices=sorted(AIRLINE_CONFIGS), help="Aerolinea para la que se arma el reporte.")
    parser.add_argument("--input", default="data/archivo original.xlsx", help="Ruta al export del sistema (archivo original).")
    parser.add_argument("--output", default=None, help="Ruta del reporte a generar. Por defecto: output/<airline>.xlsx")
    args = parser.parse_args()

    output_path = args.output or f"output/{args.airline}.xlsx"
    df = load_original(args.input)

    if AIRLINE_CONFIGS[args.airline].get("flow") == FLOW_LIQUIDACION:
        detalle = build_latam_detalle(df)
        resumen = build_latam_resumen(detalle)
        write_liquidacion(detalle, resumen, output_path)

        print(f"Liquidación generada en: {output_path}")
        print(f"  - Detalle de Facturación: {len(detalle)} filas")
        print(f"  - TOTAL LA (sin IVA): {resumen['total_la']:,.2f}")
        print(f"  - TOTAL 4M (con IVA): {resumen['total_4m']:,.2f}")
        print(f"  - TOTAL PERIODO: {resumen['total_periodo']:,.2f}")
        return

    sheets = build_airline_report(df, args.airline)
    write_report(sheets, output_path)

    print(f"Reporte generado en: {output_path}")
    for sheet_name, sheet_df in sheets.items():
        print(f"  - {sheet_name}: {len(sheet_df)} filas")


if __name__ == "__main__":
    main()
