"""CLI para generar el reporte filtrado de una aerolinea.

Uso:
    python -m src.main --airline avianca
    python -m src.main --airline gol --input "data/archivo original.xlsx" --output "output/gol.xlsx"
    python -m src.main --airline latam
    python -m src.main --airline jetsmart --input "LIQUIDACION ECS 07-2026.xlsx"         --anterior "LIQUIDACION ECS 06-2026.xlsx" [--siguiente export_agosto.xlsx]         --tc 1485 --vuelos EZE=56 MDZ=38
"""

import argparse

from src.config import AIRLINE_CONFIGS, FLOW_JETSMART, FLOW_LIQUIDACION
from src.jetsmart_builder import (
    build_jetsmart_guias,
    build_jetsmart_resumen,
    detect_jetsmart_period,
    load_jetsmart_export,
    write_jetsmart_liquidacion,
)
from src.jetsmart_cruce import cruzar_jetsmart, load_ariel
from src.liquidacion_builder import build_latam_detalle, build_latam_resumen, write_liquidacion
from src.loader import load_original
from src.report_builder import build_airline_report, write_report


def main() -> None:
    parser = argparse.ArgumentParser(description="Genera el reporte filtrado de una aerolinea a partir del export del sistema.")
    parser.add_argument("--airline", required=True, choices=sorted(AIRLINE_CONFIGS), help="Aerolinea para la que se arma el reporte.")
    parser.add_argument("--input", default="data/archivo original.xlsx", help="Ruta al export del sistema (archivo original).")
    parser.add_argument("--output", default=None, help="Ruta del reporte a generar. Por defecto: output/<airline>.xlsx")
    parser.add_argument("--tc", type=float, default=None, help="Solo JetSmart: tipo de cambio del periodo (dato manual).")
    parser.add_argument("--vuelos", nargs="*", default=[], metavar="ESTACION=N", help="Solo JetSmart: vuelos internacionales por estacion, ej. EZE=56 MDZ=38.")
    parser.add_argument("--anterior", default=None, help="Solo JetSmart: archivo de trabajo del mes anterior (hojas BD y ARIEL).")
    parser.add_argument("--siguiente", default=None, help="Solo JetSmart (opcional): export de guias del mes siguiente.")
    args = parser.parse_args()

    output_path = args.output or f"output/{args.airline}.xlsx"

    if AIRLINE_CONFIGS[args.airline].get("flow") == FLOW_JETSMART:
        if args.tc is None:
            parser.error("--tc es obligatorio para jetsmart")
        if args.anterior is None:
            parser.error("--anterior es obligatorio para jetsmart (archivo de trabajo del mes anterior)")
        vuelos_inter = {k.upper(): int(v) for k, v in (item.split("=", 1) for item in args.vuelos)}
        export = load_jetsmart_export(args.input)
        period, _ = detect_jetsmart_period(export)
        ariel, _ = load_ariel(args.input)
        ariel_anterior, _ = load_ariel(args.anterior)
        export_siguiente = load_jetsmart_export(args.siguiente) if args.siguiente else None
        cruce = cruzar_jetsmart(export, ariel, period, load_jetsmart_export(args.anterior), ariel_anterior, export_siguiente)
        guias = build_jetsmart_guias(cruce.incluidas)
        resumen = build_jetsmart_resumen(guias, args.tc, vuelos_inter)
        write_jetsmart_liquidacion(guias, resumen, period, output_path, cruce={
            "excepciones": cruce.excepciones, "incluidas_ultimo_dia": cruce.incluidas_ultimo_dia,
        })
        print(f"Cruce con Ariel: {cruce.resumen['incluidas']} guias incluidas, {cruce.resumen['excepciones']} excepciones "
              f"({cruce.resumen['excepciones_por_tipo']}), {len(cruce.incluidas_ultimo_dia)} por regla de ultimo dia.")

        print(f"Liquidación generada en: {output_path}")
        print(f"  - GUIAS: {len(guias)} filas")
        print(f"  - Ventas Netas: {resumen['ventas_netas']:,.2f}")
        print(f"  - Total a entregar a WCS: {resumen['total_wcs']:,.2f}")
        return
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
