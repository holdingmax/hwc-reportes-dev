"""Interfaz web simple para generar los reportes por aerolinea.

Es solo una capa de UI: toda la logica de carga y filtrado sigue viviendo en
src/ (loader.py, report_builder.py, config.py) y no se modifica aca. Lo unico
que se reusa de mas es AIRLINE_CONFIGS/CHARGE_TYPES para poder agrupar el
resultado por tipo de cargo al mostrarlo, sin reimplementar el filtrado.

El estilo (paleta corporativa navy + dorado, cards, tipografia) se inyecta
como CSS via st.markdown(unsafe_allow_html=True), ya que Streamlit no
permite theming tan especifico de forma nativa. Los selectores CSS estan
verificados contra el DOM real que genera Streamlit 1.58 (data-testid y la
clase "st-key-<key>" que expone st.container(key=...) para estilar
contenedores puntuales).
"""

import io

import streamlit as st

from src.config import AIRLINE_CONFIGS, CHARGE_TYPES, FLOW_LIQUIDACION
from src.liquidacion_builder import (
    build_latam_detalle,
    build_latam_resumen,
    detect_unconfirmed_station_activity,
    write_liquidacion,
)
from src.loader import load_original
from src.report_builder import build_airline_report, detect_unconfirmed_activity, write_report

CHARGE_TYPE_LABELS = {
    "delivery_fee": ("📦", "Delivery Fee"),
    "trans_electronica": ("📡", "Transmisión Electrónica"),
    "collect": ("💰", "Collect"),
}

st.set_page_config(page_title="Reportes HWC", page_icon="📄", layout="centered")

# ---------------------------------------------------------------------------
# Estilo corporativo (navy + dorado). Ver docstring del modulo.
# ---------------------------------------------------------------------------
st.markdown(
    """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

:root {
    --hwc-navy: #0B2545;
    --hwc-navy-light: #14335E;
    --hwc-gold: #C9A227;
    --hwc-gold-dark: #A6841A;
    --hwc-bg: #F4F6F9;
    --hwc-card: #FFFFFF;
    --hwc-text: #1F2933;
    --hwc-text-muted: #6B7280;
    --hwc-border: #E4E9F0;
    --hwc-success: #1E8E5A;
}

html, body, [data-testid="stApp"], [data-testid="stAppViewContainer"], [data-testid="stMain"] {
    background: var(--hwc-bg) !important;
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
    color: var(--hwc-text);
}

/* Limpiar el chrome default de Streamlit (Deploy / menu) para look interno */
[data-testid="stToolbar"] { visibility: hidden; }
[data-testid="stHeader"] { background: transparent; }

[data-testid="stMainBlockContainer"] {
    max-width: 760px;
    padding-top: 2.2rem;
    padding-bottom: 3rem;
}

h1, h2, h3 { font-family: 'Inter', sans-serif; font-weight: 800; color: var(--hwc-navy); letter-spacing: -0.01em; }
/* Nota: no pisar el color de todos los <p> de stMarkdownContainer aca:
   Streamlit tambien usa ese mismo contenedor para el texto de los botones,
   asi que una regla global de color rompe el texto blanco de los botones
   y del banner. El color de cuerpo normal ya se hereda de html/body. */
[data-testid="stCaptionContainer"] { color: var(--hwc-text-muted) !important; }
[data-testid="stWidgetLabel"] p { font-weight: 600; color: var(--hwc-text); font-size: 0.85rem; }

/* --- Titulo principal, por encima del banner de marca --- */
.hwc-page-title {
    display: flex; align-items: center; gap: 0.7rem;
    font-size: 2.6rem; font-weight: 800; color: var(--hwc-navy);
    letter-spacing: -0.02em; line-height: 1.1;
    margin: 0 0 1.2rem 0;
}
.hwc-page-title-icon { font-size: 2.3rem; line-height: 1; }

/* --- Banner superior de marca --- */
.hwc-hero {
    background: linear-gradient(135deg, var(--hwc-navy) 0%, var(--hwc-navy-light) 100%);
    border-radius: 14px;
    padding: 1.8rem 2.1rem;
    margin-bottom: 1.6rem;
    box-shadow: 0 10px 30px -14px rgba(11, 37, 69, 0.55);
    position: relative;
    overflow: hidden;
}
.hwc-hero::after {
    content: "";
    position: absolute;
    right: -50px;
    top: -60px;
    width: 180px;
    height: 180px;
    background: rgba(201, 162, 39, 0.15);
    border-radius: 50%;
}
.hwc-hero-sub { color: rgba(255, 255, 255, 0.75) !important; font-size: 0.85rem; margin: 0; position: relative; }
.hwc-hero-tag {
    display: inline-block; margin-top: 0.9rem;
    background: rgba(201, 162, 39, 0.16); border: 1px solid rgba(201, 162, 39, 0.45);
    color: var(--hwc-gold); font-size: 0.7rem; font-weight: 700; letter-spacing: 0.06em;
    text-transform: uppercase; padding: 0.28rem 0.75rem; border-radius: 100px; position: relative;
}

/* --- Cards de seccion (subir archivo / aerolinea / resultado) --- */
.st-key-card_upload[data-testid="stVerticalBlock"],
.st-key-card_config[data-testid="stVerticalBlock"],
.st-key-card_result[data-testid="stVerticalBlock"] {
    background: var(--hwc-card);
    border: 1px solid var(--hwc-border);
    border-radius: 14px;
    padding: 1.5rem 1.6rem 1.3rem 1.6rem;
    box-shadow: 0 4px 18px -10px rgba(16, 24, 40, 0.12);
    margin-bottom: 1.3rem;
}

.hwc-step {
    display: flex; align-items: center; gap: 0.6rem;
    font-weight: 700; font-size: 1.02rem; color: var(--hwc-navy);
    margin-bottom: 0.9rem;
}
.hwc-step-num {
    display: inline-flex; align-items: center; justify-content: center;
    width: 24px; height: 24px; border-radius: 50%; flex-shrink: 0;
    background: var(--hwc-gold); color: #fff; font-size: 0.8rem; font-weight: 800;
}

/* --- Dropzone de archivos --- */
[data-testid="stFileUploaderDropzone"] {
    background: #FAFBFC !important;
    border: 1.5px dashed var(--hwc-border) !important;
    border-radius: 10px !important;
}

/* --- Select --- */
div[data-baseweb="select"] > div { border-radius: 8px !important; }

/* --- Botones primarios (Generar reporte) --- */
[data-testid="stBaseButton-primary"] {
    background: linear-gradient(135deg, var(--hwc-navy) 0%, var(--hwc-navy-light) 100%) !important;
    color: #fff !important;
    border: none !important;
    border-radius: 8px !important;
    font-weight: 600 !important;
    padding: 0.55rem 1.5rem !important;
    box-shadow: 0 4px 14px -4px rgba(11, 37, 69, 0.45);
    transition: filter .15s ease, box-shadow .15s ease;
}
[data-testid="stBaseButton-primary"]:hover {
    filter: brightness(1.15);
    box-shadow: 0 6px 18px -4px rgba(11, 37, 69, 0.6);
}
[data-testid="stBaseButton-primary"] p { color: #fff !important; font-weight: 600 !important; }
[data-testid="stBaseButton-primary"]:disabled {
    background: #C6CDD6 !important;
    color: #8A94A3 !important;
    box-shadow: none;
}
[data-testid="stBaseButton-primary"]:disabled p { color: #8A94A3 !important; }
[data-testid="stBaseButton-secondary"] {
    border-radius: 8px !important;
    border: 1.5px solid var(--hwc-navy) !important;
    color: var(--hwc-navy) !important;
    font-weight: 600 !important;
}
[data-testid="stBaseButton-secondary"] p { color: var(--hwc-navy) !important; font-weight: 600 !important; }

/* El boton de descarga vive dentro del card de resultado: lo diferenciamos
   en dorado (accion final) vs. el navy de "Generar reporte". */
.st-key-card_result [data-testid="stBaseButton-primary"] {
    background: linear-gradient(135deg, var(--hwc-gold) 0%, var(--hwc-gold-dark) 100%) !important;
    box-shadow: 0 4px 14px -4px rgba(166, 132, 26, 0.5);
}
.st-key-card_result [data-testid="stBaseButton-primary"] p { color: var(--hwc-navy) !important; }

[data-testid="stAlert"] { border-radius: 10px; font-size: 0.85rem; }

/* --- Mini-cards por tipo de cargo dentro del resultado --- */
.hwc-group-card {
    background: #FAFBFC;
    border: 1px solid var(--hwc-border);
    border-radius: 10px;
    padding: 0.9rem 1.1rem;
    margin-bottom: 0.9rem;
}
.hwc-group-title {
    font-weight: 700; color: var(--hwc-navy); font-size: 0.95rem;
    margin-bottom: 0.5rem; display: flex; align-items: center; gap: 0.4rem;
}
.hwc-row { display: flex; align-items: center; gap: 0.55rem; font-size: 0.86rem; padding: 0.22rem 0; }
.hwc-dot {
    display: inline-flex; align-items: center; justify-content: center;
    width: 18px; height: 18px; border-radius: 50%; font-size: 0.65rem;
    font-weight: 800; flex-shrink: 0;
}
.hwc-dot-active { background: var(--hwc-success); color: #fff; }
.hwc-dot-empty { background: #E4E9F0; color: #9AA5B1; }
.hwc-row-active { color: var(--hwc-text); }
.hwc-row-empty { color: var(--hwc-text-muted); }
.hwc-count {
    margin-left: auto; font-weight: 700; color: var(--hwc-navy);
    background: #EEF2F8; padding: 0.1rem 0.6rem; border-radius: 100px; font-size: 0.78rem;
}

/* --- Footer --- */
.hwc-footer {
    text-align: center; color: var(--hwc-text-muted); font-size: 0.78rem;
    margin-top: 1.6rem; padding-top: 1rem; border-top: 1px solid var(--hwc-border);
}
</style>
""",
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# Titulo principal (jerarquia por encima del banner de marca, sin tocarlo)
# ---------------------------------------------------------------------------
st.markdown(
    '<div class="hwc-page-title"><span class="hwc-page-title-icon">✈️</span>Handyway Cargo</div>',
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# Banner de marca
# ---------------------------------------------------------------------------
st.markdown(
    """
<div class="hwc-hero">
    <p class="hwc-hero-sub">Generá reportes automáticos por aerolínea a partir del archivo original de Handyway Cargo.</p>
    <span class="hwc-hero-tag">Automatización de reportes</span>
</div>
""",
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# Seccion 1: subir archivo
# ---------------------------------------------------------------------------
with st.container(border=True, key="card_upload"):
    st.markdown('<div class="hwc-step"><span class="hwc-step-num">1</span>📄 Subí el archivo</div>', unsafe_allow_html=True)
    uploaded_file = st.file_uploader("Archivo original (export del sistema)", type="xlsx", label_visibility="collapsed")

# ---------------------------------------------------------------------------
# Seccion 2: elegir aerolinea + generar
# ---------------------------------------------------------------------------
with st.container(border=True, key="card_config"):
    st.markdown('<div class="hwc-step"><span class="hwc-step-num">2</span>✈️ Elegí la aerolínea</div>', unsafe_allow_html=True)
    airline_key = st.selectbox("Aerolínea", options=sorted(AIRLINE_CONFIGS), format_func=str.upper, label_visibility="collapsed")
    generate = st.button("Generar reporte", disabled=uploaded_file is None, type="primary")


def _sheets_for_charge_type(sheets: dict, charge_type_key: str) -> dict:
    """Agrupa las hojas ya generadas segun a que tipo de cargo pertenecen.

    Es una funcion puramente de presentacion: usa la misma config de src/
    para saber que prefijo/nombre de hoja corresponde a cada tipo de cargo,
    pero no vuelve a filtrar ni calcular nada.
    """
    charge_cfg = CHARGE_TYPES[charge_type_key]
    if charge_cfg.get("per_station"):
        prefix = charge_cfg["sheet_prefix"] + " "
        return {name: d for name, d in sheets.items() if name.startswith(prefix)}
    name = charge_cfg["sheet_name"]
    return {name: sheets[name]} if name in sheets else {}


def _money(value: float) -> str:
    return f"$ {value:,.2f}"


def _render_liquidacion_result(uploaded_file) -> tuple[object, dict]:
    """Corre el flujo de liquidacion de LATAM y muestra su propio resumen.

    A diferencia del reporte simple (conteo de filas por hoja), acá lo que
    importa mostrar son los totales del Resumen Facturación.
    """
    with st.spinner("Generando liquidación..."):
        df = load_original(uploaded_file)
        detalle = build_latam_detalle(df)
        resumen = build_latam_resumen(detalle)

        buffer = io.BytesIO()
        write_liquidacion(detalle, resumen, buffer)
        buffer.seek(0)

    st.success("Liquidación generada correctamente.")

    unconfirmed_stations = AIRLINE_CONFIGS["latam"].get("unconfirmed_stations", [])
    if unconfirmed_stations:
        activity = detect_unconfirmed_station_activity(df, unconfirmed_stations)
        if activity:
            detalle_txt = ", ".join(f"{station} ({_money(info['total'])})" for station, info in activity.items())
            st.warning(
                f"⚠️ Se detectaron movimientos de LATAM sin incluir en esta liquidación "
                f"(estación no validada todavía): {detalle_txt}. Ese monto queda fuera de "
                f"TOTAL PERIODO — no se está facturando."
            )

    st.markdown(
        f'<div class="hwc-group-card">'
        f'<div class="hwc-group-title">📑 Detalle de Facturación</div>'
        f'<div class="hwc-row hwc-row-active"><span class="hwc-dot hwc-dot-active">✓</span>'
        f'Guías con cargo (EZE)<span class="hwc-count">{len(detalle)} filas</span></div>'
        f'</div>',
        unsafe_allow_html=True,
    )

    st.markdown(
        f'<div class="hwc-group-card">'
        f'<div class="hwc-group-title">🧾 Resumen Facturación</div>'
        f'<div class="hwc-row hwc-row-active"><span class="hwc-dot hwc-dot-active">✓</span>'
        f'TOTAL LA (sin IVA)<span class="hwc-count">{_money(resumen["total_la"])}</span></div>'
        f'<div class="hwc-row hwc-row-active"><span class="hwc-dot hwc-dot-active">✓</span>'
        f'TOTAL 4M (con IVA)<span class="hwc-count">{_money(resumen["total_4m"])}</span></div>'
        f'<div class="hwc-row hwc-row-active"><span class="hwc-dot hwc-dot-active">✓</span>'
        f'TOTAL PERIODO<span class="hwc-count">{_money(resumen["total_periodo"])}</span></div>'
        f'</div>',
        unsafe_allow_html=True,
    )

    st.markdown(
        f'<div class="hwc-group-card">'
        f'<div class="hwc-group-title">📝 Compensación</div>'
        f'<div class="hwc-row hwc-row-empty"><span class="hwc-dot hwc-dot-empty">–</span>'
        f'pendiente de carga manual para este período</div>'
        f'</div>',
        unsafe_allow_html=True,
    )

    return buffer, resumen


# ---------------------------------------------------------------------------
# Seccion 3: resultado
# ---------------------------------------------------------------------------
if generate:
    with st.container(border=True, key="card_result"):
        st.markdown('<div class="hwc-step"><span class="hwc-step-num">3</span>📊 Resultado</div>', unsafe_allow_html=True)

        airline_cfg = AIRLINE_CONFIGS[airline_key]

        if airline_cfg.get("flow") == FLOW_LIQUIDACION:
            buffer, _ = _render_liquidacion_result(uploaded_file)
        else:
            with st.spinner("Generando reporte..."):
                df = load_original(uploaded_file)
                sheets = build_airline_report(df, airline_key)

                buffer = io.BytesIO()
                write_report(sheets, buffer)
                buffer.seek(0)

            st.success("Reporte generado correctamente.")

            unconfirmed_stations = airline_cfg.get("unconfirmed_stations", [])
            if unconfirmed_stations:
                activity = detect_unconfirmed_activity(sheets, unconfirmed_stations)
                if activity:
                    estaciones = ", ".join(activity)
                    st.warning(
                        f"⚠️ Se detectaron movimientos en {estaciones} para {airline_key.upper()}. "
                        f"La tarifa aplicada ahí todavía no está confirmada con el cliente — "
                        f"revisá los montos manualmente antes de usarlos."
                    )

            for charge_type_key in airline_cfg["charge_types"]:
                icon, label = CHARGE_TYPE_LABELS.get(charge_type_key, ("📁", charge_type_key))
                group_sheets = _sheets_for_charge_type(sheets, charge_type_key)

                rows_html = ""
                for sheet_name, sheet_df in group_sheets.items():
                    n_rows = len(sheet_df)
                    if n_rows == 0:
                        rows_html += (
                            f'<div class="hwc-row hwc-row-empty">'
                            f'<span class="hwc-dot hwc-dot-empty">–</span>{sheet_name} · sin movimiento</div>'
                        )
                    else:
                        rows_html += (
                            f'<div class="hwc-row hwc-row-active">'
                            f'<span class="hwc-dot hwc-dot-active">✓</span>{sheet_name}'
                            f'<span class="hwc-count">{n_rows} filas</span></div>'
                        )

                st.markdown(
                    f'<div class="hwc-group-card">'
                    f'<div class="hwc-group-title">{icon} {label}</div>'
                    f'{rows_html}</div>',
                    unsafe_allow_html=True,
                )

        st.download_button(
            label="⬇️ Descargar reporte",
            data=buffer,
            file_name=f"{airline_key}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

st.markdown('<div class="hwc-footer">Handyway Cargo · Automatización de reportes</div>', unsafe_allow_html=True)
