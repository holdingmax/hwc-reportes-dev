"""Interfaz web simple para generar los reportes por aerolinea.

Es solo una capa de UI: toda la logica de carga y filtrado sigue viviendo en
src/ (loader.py, report_builder.py, config.py) y no se modifica aca. Lo unico
que se reusa de mas es AIRLINE_CONFIGS/CHARGE_TYPES para poder agrupar el
resultado por tipo de cargo al mostrarlo, sin reimplementar el filtrado.

Ademas de generar el Excel como siempre, cada reporte generado se persiste en
Postgres via src/db.py (guardar_reporte_simple / guardar_liquidacion_latam) --
ver ese modulo para el detalle. Es un efecto secundario que nunca bloquea ni
cambia el Excel: si DATABASE_URL no esta configurada o la base falla, esas
funciones no hacen nada y el flujo de descarga sigue exactamente igual.

La pagina tiene dos secciones separadas: "Cargar archivo nuevo" (formulario +
resultado inmediato de procesarlo, un solo flujo) e "Historial de
Liquidaciones", de SOLO LECTURA contra lo ya persistido via
obtener_filtros_historial/obtener_historial en src/db.py, que arranca
colapsada en un expander para no mezclar el dashboard historico con lo que
se acaba de procesar. Mismo criterio fail-soft: si la base no responde,
muestra un mensaje en vez de romper la pantalla. No se usa st.tabs() a
proposito: ejecuta el codigo de todas las pestanas en cada rerun (ver nota
en _get_connection de src/db.py).

El estilo (paleta de azules extraida del logo, cards, tipografia, marca de
agua) se inyecta como CSS via st.markdown(unsafe_allow_html=True), ya que
Streamlit no permite theming tan especifico de forma nativa. Los selectores
CSS estan verificados contra el DOM real que genera Streamlit 1.58
(data-testid y la clase "st-key-<key>" que expone st.container(key=...)
para estilar contenedores puntuales). El logo (static/images/handyway-logo.png)
se embebe como data URI base64 -- no como archivo estatico servido por
separado -- para no depender de configuracion de static file serving en
Render.
"""

import base64
import io
import os

import altair as alt
import pandas as pd
import streamlit as st

from src.config import (
    AIRLINE_CONFIGS,
    CHARGE_TYPES,
    COL_COD_VUELO,
    FLOW_JETSMART,
    FLOW_LIQUIDACION,
    JETSMART_COMISION_INTER_USD_POR_VUELO,
    LATAM_SUBFACTURA_4M,
    LATAM_SUBFACTURA_LA,
)
from src.db import (
    guardar_liquidacion_jetsmart,
    guardar_liquidacion_latam,
    guardar_reporte_simple,
    obtener_detalle_totales,
    obtener_filtros_historial,
    obtener_historial,
    obtener_movimientos_de_carga,
)
from src.liquidacion_builder import (
    build_latam_detalle,
    build_latam_resumen,
    detect_latam_period_exclusions,
    detect_unconfirmed_station_activity,
    write_liquidacion,
)
from src.jetsmart_cruce import cruzar_jetsmart, has_ariel_sheet, load_ariel
from src.jetsmart_builder import (
    build_jetsmart_guias,
    build_jetsmart_resumen,
    detect_jetsmart_period,
    load_jetsmart_export,
    unconfirmed_tarifa_activity,
    write_jetsmart_liquidacion,
)
from src.loader import load_original
from src.period_utils import extract_period, period_label, period_slug
from src.report_builder import (
    build_airline_report,
    detect_period_exclusions,
    detect_unconfirmed_activity,
    write_report,
)

CHARGE_TYPE_LABELS = {
    "delivery_fee": ("📦", "Delivery Fee"),
    "trans_electronica": ("📡", "Transmisión Electrónica"),
    "collect": ("💰", "Collect"),
    "latam_liquidacion": ("🧾", "Liquidación LATAM"),
    "jetsmart_liquidacion": ("🧾", "Liquidación JetSmart"),
}

LOGO_PATH = os.path.join("static", "images", "handyway-logo.png")


def _logo_base64() -> str:
    """Lee el logo del disco y lo devuelve como data URI base64.

    Se embebe inline en el HTML (en vez de servirlo como archivo estatico)
    para no depender de que Render sirva la carpeta static/ -- funciona
    igual en cualquier hosting, sin configuracion adicional.
    """
    with open(LOGO_PATH, "rb") as f:
        encoded = base64.b64encode(f.read()).decode("ascii")
    return f"data:image/png;base64,{encoded}"


LOGO_DATA_URI = _logo_base64()

# layout="wide" a proposito: el Historial (tabla + dos graficos) necesita
# mas ancho que el formulario de carga. El ancho de cada zona lo controla el
# CSS de abajo (.st-key-zona_carga angosta, .st-key-zona_historial ancha),
# no Streamlit -- ver "Anchos de la pagina".
st.set_page_config(page_title="Reportes HWC", page_icon=LOGO_PATH, layout="wide")

# ---------------------------------------------------------------------------
# Estilo corporativo. Paleta extraida directamente del logo de Handyway
# Cargo (static/images/handyway-logo.png), no inventada: --hwc-blue-light y
# --hwc-blue son los dos tonos reales del icono (muestreados con Pillow),
# --hwc-blue-text es el tono del wordmark "Handyway Cargo", y --hwc-blue-deep
# / --hwc-blue-deepest son ese mismo matiz (H=204-205) oscurecido para tener
# un navy de marca para fondos oscuros -- mismo tono, no un azul generico.
# ---------------------------------------------------------------------------
st.markdown(
    f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

:root {{
    --hwc-blue-deepest: #112E41;
    --hwc-blue-deep: #173F59;
    --hwc-blue: #2D79AB;
    --hwc-blue-light: #3895D1;
    --hwc-blue-text: #2C5877;
    --hwc-bg-1: #DDE7EE;
    --hwc-bg-2: #C9D8E3;
    --hwc-card: #FFFFFF;
    --hwc-text: #1F2933;
    --hwc-text-muted: #64748B;
    --hwc-border: #DCE7EE;
    --hwc-success: #1E8E5A;
    --hwc-warning: #C27C0E;
}}

html, body, [data-testid="stApp"], [data-testid="stAppViewContainer"], [data-testid="stMain"] {{
    background: linear-gradient(160deg, var(--hwc-bg-1) 0%, var(--hwc-bg-2) 100%) !important;
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
    color: var(--hwc-text);
}}

/* Limpiar el chrome default de Streamlit (Deploy / menu) para look interno */
[data-testid="stToolbar"] {{ visibility: hidden; }}
[data-testid="stHeader"] {{ background: transparent; }}

/* --- Marca de agua: el logo, muy sutil, de fondo detras del contenido.
   Es un <div> real (ver hwc-watermark mas abajo en el markdown), no un
   ::before: en el DOM que genera Streamlit 1.58 el pseudo-elemento no
   llega a pintarse en este contenedor (probado en el navegador real -- el
   mismo data URI SI renderiza como elemento normal), asi que se opta por
   un div, que es la forma que efectivamente se ve. Vive en su propia caja
   (no hereda opacity de los hijos), asi que bajar su opacity no afecta el
   contenido real -- solo hace falta que el contenido tenga z-index por
   encima. Se ancla a una esquina (no al centro) a proposito: el contenido
   es una columna centrada (ver "Anchos de la pagina"), asi que un watermark
   centrado queda tapado por las cards en cualquier ancho de pantalla; en
   una esquina asoma en el margen en desktop y se recorta parcialmente en
   mobile, sin competir nunca con el contenido. "fixed" (no "absolute")
   para que quede anclado a la esquina del viewport tambien mientras se
   scrollea una pagina larga (no hay ancestros con transform que rompan el
   containing block de fixed -- verificado). --- */
.hwc-watermark {{
    position: fixed;
    inset: 0;
    background-image: url('{LOGO_DATA_URI}');
    background-repeat: no-repeat;
    background-position: bottom -60px right -60px;
    background-size: 520px;
    opacity: 0.07;
    pointer-events: none;
    z-index: 0;
}}
/* --- Anchos de la pagina. La pagina usa layout="wide" y acota cada zona:
   - el contenedor general llega hasta 1180px (encabezado e Historial);
   - .st-key-zona_carga (formulario de carga + resultado) queda en 760px,
     centrado: un formulario de un solo paso a la vez se lee mejor angosto;
   - .st-key-zona_historial usa el ancho completo, para que la tabla y los
     dos graficos lado a lado no queden cortados.
   Las clases st-key-<key> las genera Streamlit 1.58 para st.container(key=...)
   (verificado contra el DOM real). En mobile todo colapsa al 100%. --- */
[data-testid="stMainBlockContainer"] {{
    max-width: 1180px;
    margin: 0 auto;
    padding: 2.2rem 2rem 3rem 2rem;
    position: relative;
    z-index: 1;
}}
.st-key-zona_carga {{
    max-width: 760px;
    width: 100%;
    margin-left: auto;
    margin-right: auto;
}}
.st-key-zona_historial {{ width: 100%; }}
/* Dentro de la zona de carga las cards ya quedan separadas por el gap
   propio de Streamlit: sin margen extra, para que los pasos se lean como
   un solo flujo y no como bloques sueltos. */
.st-key-zona_carga [data-testid="stVerticalBlock"][class*="st-key-card_"] {{ margin-bottom: 0; }}

h1, h2, h3 {{ font-family: 'Inter', sans-serif; font-weight: 800; color: var(--hwc-blue-text); letter-spacing: -0.01em; }}
/* Nota: no pisar el color de todos los <p> de stMarkdownContainer aca:
   Streamlit tambien usa ese mismo contenedor para el texto de los botones,
   asi que una regla global de color rompe el texto blanco de los botones
   y del banner. El color de cuerpo normal ya se hereda de html/body. */
[data-testid="stCaptionContainer"] {{ color: var(--hwc-text-muted) !important; }}
[data-testid="stWidgetLabel"] p {{ font-weight: 600; color: var(--hwc-text); font-size: 0.85rem; }}

/* --- Encabezado: un unico banner de marca, compacto (logo sobre chip
   blanco -- el logo es azul y no contrasta sobre el degrade oscuro). --- */
.hwc-hero {{
    background: linear-gradient(135deg, var(--hwc-blue-deep) 0%, var(--hwc-blue) 100%);
    border-radius: 16px;
    padding: 1.15rem 1.6rem;
    margin-bottom: 0.4rem;
    box-shadow: 0 10px 30px -14px rgba(17, 46, 65, 0.55);
    position: relative;
    overflow: hidden;
}}
.hwc-hero::after {{
    content: "";
    position: absolute;
    right: -50px;
    top: -60px;
    width: 180px;
    height: 180px;
    background: rgba(255, 255, 255, 0.09);
    border-radius: 50%;
}}
.hwc-hero-brand {{ display: flex; align-items: center; gap: 1rem; position: relative; }}
.hwc-hero-logo {{
    display: inline-flex; align-items: center; justify-content: center; flex-shrink: 0;
    width: 52px; height: 52px; border-radius: 12px; background: #fff;
    box-shadow: 0 4px 12px -6px rgba(17, 46, 65, 0.6);
}}
.hwc-hero-logo img {{ width: 38px; height: auto; }}
.hwc-hero-title {{ color: #fff; font-size: 1.45rem; font-weight: 800; letter-spacing: -0.01em; line-height: 1.15; }}
.hwc-hero-sub {{ color: rgba(255, 255, 255, 0.82) !important; font-size: 0.85rem; margin: 0.15rem 0 0 0; }}

/* --- Cards de seccion (subir archivo / aerolinea / resultado / login / historial) --- */
.st-key-card_upload[data-testid="stVerticalBlock"],
.st-key-card_config[data-testid="stVerticalBlock"],
.st-key-card_datos[data-testid="stVerticalBlock"],
.st-key-card_result[data-testid="stVerticalBlock"],
.st-key-card_login[data-testid="stVerticalBlock"],
.st-key-card_historial[data-testid="stVerticalBlock"] {{
    background: var(--hwc-card);
    border: 1px solid var(--hwc-border);
    border-radius: 16px;
    padding: 1.7rem 1.8rem 1.5rem 1.8rem;
    box-shadow: 0 4px 22px -10px rgba(23, 63, 89, 0.14);
    margin-bottom: 1.3rem;
}}

/* --- Titulo de cada una de las dos secciones de la pagina --- */
.hwc-section-title {{
    font-size: 1.35rem; font-weight: 800; color: var(--hwc-blue-text);
    letter-spacing: -0.01em; line-height: 1.2;
    border-left: 4px solid var(--hwc-blue); padding-left: 0.7rem;
    margin: 1.6rem 0 0.25rem 0;
}}
.hwc-section-sub {{ color: var(--hwc-text-muted); font-size: 0.88rem; margin: 0 0 1rem 0.95rem; }}

.hwc-step {{
    display: flex; align-items: flex-start; gap: 0.75rem;
    font-weight: 700; font-size: 1.05rem; color: var(--hwc-blue-text);
    margin-bottom: 0.9rem;
}}
.hwc-step-title {{ line-height: 27px; }}
.hwc-step-help {{ font-weight: 400; font-size: 0.85rem; color: var(--hwc-text-muted); line-height: 1.45; margin-top: 0.1rem; }}
.hwc-step-help b {{ color: var(--hwc-text); font-weight: 600; }}
.hwc-result-title {{ font-weight: 800; font-size: 1.1rem; color: var(--hwc-blue-text); margin-bottom: 0.4rem; }}

/* --- Que falta para procesar (explica el boton deshabilitado) --- */
.hwc-req {{ background: #F5F8FB; border: 1px solid var(--hwc-border); border-radius: 12px; padding: 0.85rem 1.1rem; margin-bottom: 0.8rem; }}
.hwc-req-title {{ font-size: 0.82rem; font-weight: 700; color: var(--hwc-blue-text); margin-bottom: 0.35rem; }}
.hwc-req ul {{ list-style: none; margin: 0; padding: 0; }}
.hwc-req li {{ display: flex; align-items: center; gap: 0.55rem; font-size: 0.86rem; padding: 0.18rem 0; }}
.hwc-req-ok {{ color: var(--hwc-text-muted); text-decoration: line-through; text-decoration-color: rgba(100, 116, 139, 0.45); }}
.hwc-req-pending {{ color: var(--hwc-text); }}
.hwc-ready {{ display: flex; align-items: center; gap: 0.55rem; font-size: 0.88rem; font-weight: 600; color: var(--hwc-success); margin-bottom: 0.8rem; }}

/* --- Selector de aerolinea (segmented control): opciones parejas y grandes --- */
.st-key-aerolinea_sel button {{ min-height: 2.6rem; font-weight: 600; }}
.st-key-aerolinea_sel button p {{ font-size: 0.92rem; }}
.hwc-step-num {{
    display: inline-flex; align-items: center; justify-content: center;
    width: 27px; height: 27px; border-radius: 50%; flex-shrink: 0;
    background: linear-gradient(135deg, var(--hwc-blue-light) 0%, var(--hwc-blue) 100%);
    color: #fff; font-size: 0.85rem; font-weight: 800;
    box-shadow: 0 3px 8px -2px rgba(45, 121, 171, 0.55);
}}

/* --- Dropzone de archivos --- */
[data-testid="stFileUploaderDropzone"] {{
    background: #FAFBFD !important;
    border: 1.5px dashed var(--hwc-border) !important;
    border-radius: 10px !important;
}}

/* --- Select --- */
div[data-baseweb="select"] > div {{ border-radius: 8px !important; }}

/* --- Botones primarios (Generar reporte) --- */
[data-testid="stBaseButton-primary"] {{
    background: linear-gradient(135deg, var(--hwc-blue-deep) 0%, var(--hwc-blue) 100%) !important;
    color: #fff !important;
    border: none !important;
    border-radius: 9px !important;
    font-weight: 600 !important;
    padding: 0.6rem 1.6rem !important;
    box-shadow: 0 4px 14px -4px rgba(23, 63, 89, 0.45);
    transition: filter .15s ease, box-shadow .15s ease, transform .05s ease;
}}
[data-testid="stBaseButton-primary"]:hover {{
    filter: brightness(1.12);
    box-shadow: 0 6px 18px -4px rgba(23, 63, 89, 0.6);
}}
[data-testid="stBaseButton-primary"]:active {{ transform: translateY(1px); }}
[data-testid="stBaseButton-primary"] p {{ color: #fff !important; font-weight: 600 !important; }}
[data-testid="stBaseButton-primary"]:disabled {{
    background: #C7D2DA !important;
    color: #8A96A3 !important;
    box-shadow: none;
}}
[data-testid="stBaseButton-primary"]:disabled p {{ color: #8A96A3 !important; }}
[data-testid="stBaseButton-secondary"] {{
    border-radius: 9px !important;
    border: 1.5px solid var(--hwc-blue) !important;
    color: var(--hwc-blue-text) !important;
    font-weight: 600 !important;
    transition: background-color .15s ease;
}}
[data-testid="stBaseButton-secondary"] p {{ color: var(--hwc-blue-text) !important; font-weight: 600 !important; }}
[data-testid="stBaseButton-secondary"]:hover {{ background: var(--hwc-bg-1) !important; }}

/* El boton de descarga vive dentro del card de resultado: lo diferenciamos
   con el azul claro de acento (accion final) vs. el azul profundo de
   "Generar reporte". */
.st-key-card_result [data-testid="stBaseButton-primary"] {{
    background: linear-gradient(135deg, var(--hwc-blue-light) 0%, var(--hwc-blue) 100%) !important;
    box-shadow: 0 4px 14px -4px rgba(56, 149, 209, 0.55);
}}

[data-testid="stAlert"] {{ border-radius: 12px; font-size: 0.85rem; padding: 0.85rem 1rem; }}

/* --- Mini-cards por tipo de cargo dentro del resultado --- */
.hwc-group-card {{
    background: #FAFBFD;
    border: 1px solid var(--hwc-border);
    border-radius: 12px;
    padding: 1rem 1.15rem;
    margin-bottom: 0.9rem;
}}
.hwc-group-title {{
    font-weight: 700; color: var(--hwc-blue-text); font-size: 0.95rem;
    margin-bottom: 0.55rem; display: flex; align-items: center; gap: 0.4rem;
}}
.hwc-row {{ display: flex; align-items: center; gap: 0.55rem; font-size: 0.86rem; padding: 0.24rem 0; }}
.hwc-dot {{
    display: inline-flex; align-items: center; justify-content: center;
    width: 18px; height: 18px; border-radius: 50%; font-size: 0.65rem;
    font-weight: 800; flex-shrink: 0;
}}
.hwc-dot-active {{ background: var(--hwc-success); color: #fff; }}
.hwc-dot-success {{ background: var(--hwc-success); color: #fff; }}
.hwc-dot-pending {{ background: #fff; border: 1.5px solid #B8C4CF; }}
/* Estado "revisar": ambar, reservado para excepciones/avisos (siempre con "!"). */
.hwc-dot-warn {{ background: var(--hwc-warning); color: #fff; }}

/* --- Renglones de las cards de resumen (ver _resumen_card) --- */
.hwc-line {{
    display: flex; align-items: center; justify-content: space-between; gap: 1rem;
    font-size: 0.87rem; color: var(--hwc-text); padding: 0.36rem 0;
    border-bottom: 1px solid #EEF3F7;
}}
.hwc-line:last-child {{ border-bottom: none; }}
.hwc-line-label {{ display: flex; align-items: center; gap: 0.55rem; }}
.hwc-line-value {{ font-weight: 600; white-space: nowrap; font-variant-numeric: tabular-nums; color: var(--hwc-text); }}
.hwc-line-subtotal {{ font-weight: 700; }}
.hwc-line-subtotal .hwc-line-value {{ font-weight: 700; }}
.hwc-line-total {{ border-top: 1.5px solid var(--hwc-blue-text); border-bottom: none; margin-top: 0.3rem; padding-top: 0.6rem; font-weight: 800; color: var(--hwc-blue-text); }}
.hwc-line-total .hwc-line-value {{ font-weight: 800; color: var(--hwc-blue-text); background: #EAF1F6; padding: 0.2rem 0.7rem; border-radius: 100px; }}
.hwc-line-vacio {{ color: var(--hwc-text-muted); }}
.hwc-line-warn .hwc-line-value {{ color: #8A5A12; background: #FBF0DC; padding: 0.1rem 0.6rem; border-radius: 100px; }}
.hwc-line-group {{
    font-size: 0.74rem; font-weight: 700; letter-spacing: 0.05em; text-transform: uppercase;
    color: var(--hwc-text-muted); padding: 0.7rem 0 0.2rem 0;
}}
.hwc-line-group:first-child {{ padding-top: 0; }}
.hwc-dot-empty {{ background: #E4E9F0; color: #9AA5B1; }}
.hwc-row-active {{ color: var(--hwc-text); }}
.hwc-row-empty {{ color: var(--hwc-text-muted); }}
.hwc-count {{
    margin-left: auto; font-weight: 700; color: var(--hwc-blue-text);
    background: #EAF1F6; padding: 0.15rem 0.65rem; border-radius: 100px; font-size: 0.78rem;
}}

/* --- Pantalla de acceso: card angosta y centrada (con layout="wide",
   sin esto se estiraria a todo el ancho de la pagina) --- */
.st-key-card_login[data-testid="stVerticalBlock"] {{ max-width: 440px; width: 100%; margin: 9vh auto 0 auto; }}
.hwc-login-wrap {{ display: flex; flex-direction: column; align-items: center; text-align: center; }}
.hwc-login-logo {{ height: 4.5rem; width: auto; margin-bottom: 1.1rem; }}
.hwc-login-title {{ font-size: 1.3rem; font-weight: 700; color: var(--hwc-blue-text); margin-bottom: 0.3rem; }}
.hwc-login-sub {{ color: var(--hwc-text-muted); font-size: 0.85rem; margin-bottom: 1.4rem; }}
.st-key-card_login [data-testid="stTextInput"] input {{
    border-radius: 9px !important;
    text-align: center;
}}

/* --- Graficos del mini-dashboard del Historial: ocultar el menu "..."
   (View Source / View Compiled Vega / Open in Vega Editor) que vega-embed
   agrega solo por defecto -- es un menu de desarrollador, no algo para
   los usuarios finales. Verificado contra el DOM real: es un <details>
   sin clase propia, hijo directo de [data-testid="stVegaLiteChart"]. --- */
[data-testid="stVegaLiteChart"] details {{
    display: none !important;
}}

/* --- Historial --- */
/* El desplegable es la puerta de entrada: que se vea como un boton/card. */
.st-key-zona_historial [data-testid="stExpander"] details {{
    background: var(--hwc-card); border: 1px solid var(--hwc-border) !important;
    border-radius: 14px; box-shadow: 0 4px 22px -10px rgba(23, 63, 89, 0.14);
}}
.st-key-zona_historial [data-testid="stExpander"] summary {{ padding: 0.95rem 1.2rem; }}
.st-key-zona_historial [data-testid="stExpander"] summary p {{ font-weight: 700; color: var(--hwc-blue-text); font-size: 0.95rem; }}
.st-key-zona_historial [data-testid="stExpander"] summary:hover p {{ color: var(--hwc-blue); }}
.st-key-zona_historial [data-testid="stExpanderDetails"] {{ padding: 0.4rem 1.4rem 1.4rem 1.4rem; }}

.hwc-kpis {{ display: grid; grid-template-columns: 1.5fr 1fr 1fr 1fr; gap: 0.9rem; margin: 0.4rem 0 1.2rem 0; }}
.hwc-kpi {{ background: #F5F8FB; border: 1px solid var(--hwc-border); border-radius: 12px; padding: 0.9rem 1.05rem; }}
.hwc-kpi-hero {{ background: linear-gradient(135deg, var(--hwc-blue-deep) 0%, var(--hwc-blue) 100%); border: none; }}
.hwc-kpi-label {{ font-size: 0.75rem; font-weight: 700; letter-spacing: 0.04em; text-transform: uppercase; color: var(--hwc-text-muted); }}
.hwc-kpi-value {{ font-size: 1.45rem; font-weight: 800; color: var(--hwc-blue-text); margin-top: 0.2rem; line-height: 1.2; }}
.hwc-kpi-value-sm {{ font-size: 1.05rem; padding-top: 0.25rem; }}
.hwc-kpi-note {{ font-size: 0.75rem; color: var(--hwc-text-muted); margin-top: 0.25rem; }}
.hwc-kpi-hero .hwc-kpi-label, .hwc-kpi-hero .hwc-kpi-note {{ color: rgba(255, 255, 255, 0.78); }}
.hwc-kpi-hero .hwc-kpi-value {{ color: #fff; font-size: 1.7rem; }}

.hwc-chart-title {{ font-weight: 700; color: var(--hwc-blue-text); font-size: 0.95rem; margin-bottom: 0.4rem; }}
.hwc-chart-sub {{ font-weight: 500; color: var(--hwc-text-muted); font-size: 0.82rem; margin-left: 0.35rem; }}
.hwc-detail-help {{ color: var(--hwc-text-muted); font-size: 0.85rem; margin: -0.15rem 0 0.6rem 0; }}
.st-key-card_detalle[data-testid="stVerticalBlock"] {{
    background: #FAFBFD; border: 1px solid var(--hwc-border); border-radius: 14px;
    padding: 1.2rem 1.3rem; margin-top: 0.6rem;
}}

/* --- Footer --- */
.hwc-footer {{
    text-align: center; color: var(--hwc-text-muted); font-size: 0.78rem;
    margin-top: 1.6rem; padding-top: 1rem; border-top: 1px solid var(--hwc-border);
}}

/* --- Responsive: pantallas chicas --- */
@media (max-width: 480px) {{
    [data-testid="stMainBlockContainer"] {{ padding-left: 1rem !important; padding-right: 1rem !important; }}
    .hwc-hero {{ padding: 1rem 1.1rem; }}
    .hwc-hero-logo {{ width: 44px; height: 44px; }}
    .hwc-hero-logo img {{ width: 32px; }}
    .hwc-hero-title {{ font-size: 1.25rem; }}
    .st-key-card_upload[data-testid="stVerticalBlock"],
    .st-key-card_config[data-testid="stVerticalBlock"],
    .st-key-card_datos[data-testid="stVerticalBlock"],
    .st-key-card_result[data-testid="stVerticalBlock"],
    .st-key-card_login[data-testid="stVerticalBlock"],
    .st-key-card_historial[data-testid="stVerticalBlock"] {{ padding: 1.2rem 1.15rem; }}
    .hwc-watermark {{ background-size: 320px; background-position: bottom -30px right -30px; }}
    .st-key-zona_historial [data-testid="stExpanderDetails"] {{ padding: 0.3rem 0.8rem 1rem 0.8rem; }}
    .st-key-card_detalle[data-testid="stVerticalBlock"] {{ padding: 0.9rem 0.9rem; }}
}}
/* Tablet y mobile: los indicadores pasan a 2 columnas (el total, completo arriba). */
@media (max-width: 900px) {{
    .hwc-kpis {{ grid-template-columns: 1fr 1fr; }}
    .hwc-kpi-hero, .hwc-kpi:last-child {{ grid-column: 1 / -1; }}
}}
</style>
<div class="hwc-watermark"></div>
""",
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# Acceso restringido (opcional): compara contra APP_PASSWORD, una variable
# de entorno configurada en Render (Settings > Environment) -- nunca
# hardcodeada en el codigo. Si no esta configurada (ej. en local), no
# bloquea a nadie. Es una traba basica para que el link no quede abierto a
# cualquiera, no un sistema de login: no hay sesion persistente entre
# pestañas o refrescos de pagina (st.session_state vive en esa conexion).
# ---------------------------------------------------------------------------
APP_PASSWORD = os.environ.get("APP_PASSWORD")


def _check_password() -> bool:
    if not APP_PASSWORD:
        return True
    if st.session_state.get("hwc_authenticated"):
        return True

    with st.container(border=True, key="card_login"):
        st.markdown(
            f"""
            <div class="hwc-login-wrap">
                <img class="hwc-login-logo" src="{LOGO_DATA_URI}" alt="Handyway Cargo" />
                <div class="hwc-login-title">🔒 Acceso restringido</div>
                <div class="hwc-login-sub">Ingresá la contraseña para ver los reportes de Handyway Cargo.</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        password = st.text_input(
            "Contraseña", type="password", label_visibility="collapsed", placeholder="Contraseña"
        )
        if password:
            if password == APP_PASSWORD:
                st.session_state["hwc_authenticated"] = True
                st.rerun()
            else:
                st.error("Contraseña incorrecta.", icon=":material/lock:")
    return False


if not _check_password():
    st.stop()

# ---------------------------------------------------------------------------
# Funciones auxiliares de presentacion (definidas antes de usarse mas abajo,
# tanto en la seccion de carga como en Liquidaciones).
# ---------------------------------------------------------------------------

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


def _num(value: float, decimales: int = 2) -> str:
    """Numero en formato argentino: 1.234.567,89. Es el mismo formato en que
    st.dataframe muestra los numeros en un navegador en castellano, asi que
    las cards y las tablas coinciden."""
    texto = f"{value:,.{decimales}f}"
    return texto.replace(",", "").replace(".", ",").replace("", ".")


def _money(value: float) -> str:
    return f"$ {_num(value)}"


def _money_corto(value: float) -> str:
    """Monto abreviado para rotulos de graficos: $ 446,0 M."""
    if abs(value) >= 1e9:
        return f"$ {_num(value / 1e9, 1)} mil M"
    if abs(value) >= 1e6:
        return f"$ {_num(value / 1e6, 1)} M"
    return f"$ {_num(value, 0)}"


# Montos abreviados para el eje de los graficos ("$ 446,0 M"),
# como expresion de Vega -- el formato "s" de d3 usaria "G" para miles de
# millones, que no se lee en castellano.
_VEGA_MONTO_CORTO = (
    "datum.value >= 1e9 ? '$ ' + replace(format(datum.value / 1e9, '.1f'), '.', ',') + ' mil M' : "
    "datum.value >= 1e6 ? '$ ' + replace(format(datum.value / 1e6, '.1f'), '.', ',') + ' M' : "
    "'$ ' + replace(format(datum.value, ',.0f'), regexp(',', 'g'), '.')"
)
_CHART_COLOR = "#2D79AB"   # --hwc-blue: una sola serie por grafico, un solo color
_CHART_TEXT = "#1F2933"    # --hwc-text: los valores van en color de texto, no de serie
_CHART_MUTED = "#64748B"   # --hwc-text-muted
_CHART_GRID = "#E8EEF3"


def _grafico_barras(
    serie: pd.Series, nombre_categoria: str, *, horizontal: bool, orden: list[str] | None = None,
) -> alt.Chart:
    """Grafico de barras del Historial (Altair directo, no st.bar_chart,
    para controlar eje en cero, formato de moneda y orden).

    Una sola serie, un solo color de marca, sin leyenda (el titulo la
    nombra), con el valor rotulado en cada barra -- son pocas -- y tooltip
    con el monto exacto. horizontal=True para comparar categorias (se
    leen los nombres sin rotarlos); False para una evolucion en el tiempo.
    orden fuerza el orden del eje de categorias (ej. cronologico).
    """
    df = serie.reset_index()
    df.columns = [nombre_categoria, "Total facturado"]
    # Rotulo y tooltip armados aca, con el mismo formato que el resto de la app.
    df["rotulo"] = df["Total facturado"].map(_money_corto)
    df["monto"] = df["Total facturado"].map(_money)
    # Aire para el rotulo a la derecha/arriba de la barra mas larga (mas en
    # horizontal: en mobile el ancho es chico y el rotulo no se tiene que cortar).
    tope = float(df["Total facturado"].max() or 1) * (1.38 if horizontal else 1.22)
    # Orden explicito: "-x" no se respeta en un grafico de dos capas.
    sort = orden if orden is not None else df.sort_values("Total facturado", ascending=False)[nombre_categoria].tolist()
    cat_axis = alt.Axis(title=None, domain=False, ticks=False, labelPadding=8, labelColor=_CHART_TEXT,
                        labelFontSize=12, labelAngle=0)
    val_axis = alt.Axis(title=None, domain=False, ticks=False, grid=True, gridColor=_CHART_GRID,
                        labelColor=_CHART_MUTED, labelExpr=_VEGA_MONTO_CORTO, tickCount=4)
    val_scale = alt.Scale(zero=True, domain=[0, tope])
    if horizontal:
        # Sin eje de valores: cada barra ya lleva su monto rotulado, y en
        # mobile los valores del eje se pisaban entre si.
        enc = dict(y=alt.Y(f"{nombre_categoria}:N", sort=sort, axis=cat_axis),
                   x=alt.X("Total facturado:Q", scale=val_scale, axis=None))
    else:
        enc = dict(x=alt.X(f"{nombre_categoria}:N", sort=sort, axis=cat_axis),
                   y=alt.Y("Total facturado:Q", scale=val_scale, axis=val_axis))
    base = alt.Chart(df).encode(
        **enc,
        tooltip=[
            alt.Tooltip(f"{nombre_categoria}:N", title=nombre_categoria),
            alt.Tooltip("monto:N", title="Total facturado"),
        ],
    )
    barras = base.mark_bar(color=_CHART_COLOR, cornerRadiusEnd=4, size=26 if horizontal else 44)
    rotulos = base.mark_text(
        color=_CHART_TEXT, fontSize=12, fontWeight=600,
        **({"align": "left", "dx": 6} if horizontal else {"baseline": "bottom", "dy": -6}),
    ).encode(text="rotulo:N")
    alto = max(150, 48 * len(df)) if horizontal else 240
    return (barras + rotulos).properties(height=alto).configure_view(strokeWidth=0)


def _obtener_detalle_liquidacion(fila) -> tuple[pd.DataFrame, dict | None, str | None] | None:
    """Detalle linea por linea de una liquidacion del Historial, mas el
    Resumen Facturacion (con desglose de IVA) cuando aplica a LATAM.

    Se calcula con la MISMA logica que arma el Excel real
    (build_airline_report / build_latam_detalle + build_latam_resumen)
    aplicada sobre los movimientos ya guardados de esa carga -- no
    reimplementa ningun filtro ni calculo de IVA, solo reusa esas
    funciones para que lo mostrado en pantalla coincida exactamente con
    lo que trae el Excel de esa combinacion.

    Devuelve (detalle_df, resumen, sheet_name): resumen es el dict de
    build_latam_resumen() para LATAM (con total_la/neto_gravado/iva/
    total_4m/total_periodo/sums, igual que la hoja "Resumen Facturación"
    del Excel real), o None para el resto (Avianca/Gol no tienen resumen
    formal, son hojas simples). sheet_name es el nombre de hoja real
    (para reconstruir el Excel con write_report({sheet_name: detalle})) o
    None para LATAM, que no lo necesita -- write_liquidacion() ya pone su
    propio nombre de hoja internamente.

    None (no la tupla) si la base no responde o no se pudo reconstruir
    el detalle.
    """
    period = (fila["periodo_mes"], fila["periodo_anio"])
    tipo_cargo = fila["tipo_cargo"]

    if tipo_cargo == "jetsmart_liquidacion":
        # JetSmart no guarda movimientos_awb (ver guardar_liquidacion_jetsmart):
        # las guias y los datos manuales (TC, vuelos) viven en
        # detalle_totales, y el resumen se recalcula con la MISMA funcion
        # que arma el Excel real.
        detalle_totales = obtener_detalle_totales(int(fila["id"]))
        if not detalle_totales:
            return None
        guias = pd.DataFrame(detalle_totales["guias"])
        parametros = detalle_totales["parametros"]
        resumen = build_jetsmart_resumen(guias, parametros["tipo_cambio"], parametros["vuelos_inter"])
        resumen["_cruce"] = detalle_totales.get("cruce")
        return guias, resumen, None

    movimientos = obtener_movimientos_de_carga(int(fila["carga_id"]))
    if movimientos is None:
        return None

    if tipo_cargo == "latam_liquidacion":
        detalle = build_latam_detalle(movimientos, period=period)
        resumen = build_latam_resumen(detalle)
        return detalle, resumen, None

    charge_cfg = CHARGE_TYPES.get(tipo_cargo)
    if charge_cfg is None:
        return None
    sheets = build_airline_report(movimientos, fila["aerolinea"], period=period)
    if charge_cfg.get("per_station"):
        sheet_name = f"{charge_cfg['sheet_prefix']} {fila['estacion']}"
    else:
        sheet_name = charge_cfg["sheet_name"]
    detalle = sheets.get(sheet_name)
    if detalle is None:
        return None
    return detalle, None, sheet_name


# Nombre de cada aerolinea tal como se muestra en pantalla (las claves de
# AIRLINE_CONFIGS son internas, en minuscula).
AIRLINE_NAMES = {"avianca": "Avianca", "gol": "Gol", "latam": "LATAM", "jetsmart": "JetSmart"}


def _airline_name(airline_key: str) -> str:
    return AIRLINE_NAMES.get(airline_key, airline_key.upper())


# Que archivo hay que subir en el paso 2, segun la aerolinea elegida.
def _archivo_hint(airline_key: str | None) -> str:
    if airline_key is None:
        return "Primero elegí la aerolínea: según cuál sea, cambia el archivo que hay que subir."
    if AIRLINE_CONFIGS[airline_key].get("flow") == FLOW_JETSMART:
        return (
            "El <b>archivo de trabajo del mes</b> (LIQUIDACION ECS): la hoja BD con el export de guías "
            "primero y la hoja ARIEL con el archivo de Ariel."
        )
    return "El <b>export del sistema</b> (archivo original.xlsx), tal como sale, sin editar."


def _render_exito(airline_key: str, period: tuple[str, str], buffer) -> None:
    """Cierre del procesamiento: un solo mensaje de exito y la descarga.

    "Guardado en el historial" solo se afirma si la base responde: el
    guardado (src/db.py) es fail-soft y no avisa si fallo, asi que se usa
    obtener_filtros_historial() como chequeo de que la base esta
    disponible, en vez de prometer algo que puede no haber pasado.
    """
    nombre = _airline_name(airline_key)
    guardado = obtener_filtros_historial() is not None
    if guardado:
        st.success(
            f"**Listo: {nombre} · {period_label(period)}.** Quedó guardada en el Historial de Liquidaciones.",
            icon=":material/check_circle:",
        )
    else:
        st.warning(
            f"**{nombre} · {period_label(period)} generada, pero sin guardar en el historial** "
            "(la base de datos no responde). Descargá el Excel para no perderla.",
            icon=":material/cloud_off:",
        )
    st.download_button(
        label="Descargar Excel",
        icon=":material/download:",
        data=buffer,
        file_name=f"{airline_key}_{period_slug(period)}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        # Sin rerun al descargar: el resultado sigue en pantalla.
        on_click="ignore",
        key="resultado_descarga",
    )


def _resumen_card(titulo: str, filas: list[tuple]) -> str:
    """Card de resumen con un unico criterio visual en toda la app.

    filas: (tipo, etiqueta, valor). Tipos:
    - "item": renglon de un estado contable (etiqueta a la izquierda, monto
      a la derecha, sin icono: no es un estado, es un dato).
    - "grupo": encabezado de un sub-bloque (ej. "LATAM AIRLINES").
    - "subtotal" / "total": en negrita; "total" con separador arriba.
    - "ok" / "warn" / "vacio": estados, con punto de color + simbolo (nunca
      color solo): verde = salio bien, ambar = hay que revisar, gris = sin
      movimiento. El verde no se usa para datos neutros.
    """
    html = ""
    for tipo, etiqueta, valor in filas:
        valor_html = f'<span class="hwc-line-value">{valor}</span>' if valor is not None else ""
        if tipo == "grupo":
            html += f'<div class="hwc-line-group">{etiqueta}</div>'
        elif tipo in ("ok", "warn", "vacio"):
            dot = {"ok": ("hwc-dot-success", "✓"), "warn": ("hwc-dot-warn", "!"), "vacio": ("hwc-dot-empty", "–")}[tipo]
            html += (
                f'<div class="hwc-line hwc-line-{tipo}"><span class="hwc-line-label">'
                f'<span class="hwc-dot {dot[0]}">{dot[1]}</span>{etiqueta}</span>{valor_html}</div>'
            )
        else:
            html += f'<div class="hwc-line hwc-line-{tipo}"><span class="hwc-line-label">{etiqueta}</span>{valor_html}</div>'
    return f'<div class="hwc-group-card"><div class="hwc-group-title">{titulo}</div>{html}</div>'


def _render_liquidacion_result(uploaded_file) -> tuple[object, dict, tuple[str, str]]:
    """Corre el flujo de liquidacion de LATAM y muestra su propio resumen.

    A diferencia del reporte simple (conteo de filas por hoja), acá lo que
    importa mostrar son los totales del Resumen Facturación.
    """
    with st.spinner("Generando liquidación..."):
        df = load_original(uploaded_file)
        period, excluded = detect_latam_period_exclusions(df)
        detalle = build_latam_detalle(df, period=period)
        resumen = build_latam_resumen(detalle)

        buffer = io.BytesIO()
        write_liquidacion(detalle, resumen, buffer)
        buffer.seek(0)

        guardar_liquidacion_latam(
            nombre_archivo=uploaded_file.name,
            file_bytes=uploaded_file.getvalue(),
            df=df,
            period=period,
            detalle=detalle,
            resumen=resumen,
        )

    _render_exito("latam", period, buffer)
    st.caption(f"Período detectado: {period_label(period)}")

    if not excluded.empty:
        breakdown = excluded[COL_COD_VUELO].map(extract_period).value_counts()
        detalle_txt = ", ".join(f"{n} de {period_label(p)}" for p, n in breakdown.items())
        fila_word = "fila" if len(excluded) == 1 else "filas"
        st.warning(
            f"Se excluyeron {len(excluded)} {fila_word} fuera del período detectado "
            f"({period_label(period)}) de LATAM: {detalle_txt}.",
            icon=":material/warning:",
        )

    unconfirmed_stations = AIRLINE_CONFIGS["latam"].get("unconfirmed_stations", [])
    if unconfirmed_stations:
        activity = detect_unconfirmed_station_activity(df, unconfirmed_stations, period=period)
        if activity:
            detalle_txt = ", ".join(f"{station} ({_money(info['total'])})" for station, info in activity.items())
            st.warning(
                f"Se detectaron movimientos de LATAM sin incluir en esta liquidación "
                f"(estación no validada todavía): {detalle_txt}. Ese monto queda fuera de "
                f"TOTAL PERIODO — no se está facturando.",
                icon=":material/warning:",
            )

    st.markdown(_latam_resumen_html(resumen), unsafe_allow_html=True)
    st.markdown(
        _resumen_card("Detalle y compensación", [
            ("ok", "Detalle de Facturación: guías con cargo (EZE)", f"{_num(len(detalle), 0)} filas"),
            ("vacio", "Compensación: pendiente de carga manual para este período", None),
        ]),
        unsafe_allow_html=True,
    )

    return buffer, resumen, period

def _latam_resumen_html(resumen: dict) -> str:
    """Resumen Facturacion de LATAM, con el mismo desglose que la hoja real
    del Excel (build_latam_resumen, no una reimplementacion del IVA)."""
    filas = [("grupo", "LATAM AIRLINES", None)]
    filas += [("item", col, _money(resumen["sums"][col])) for col in LATAM_SUBFACTURA_LA]
    filas += [("subtotal", "TOTAL LA (sin IVA)", _money(resumen["total_la"])), ("grupo", "LAN ARGENTINA", None)]
    filas += [("item", col, _money(resumen["sums"][col])) for col in LATAM_SUBFACTURA_4M]
    filas += [
        ("item", "IVA (21%, informativo)", _money(resumen["iva"])),
        ("subtotal", "TOTAL 4M (con IVA)", _money(resumen["total_4m"])),
        ("total", "TOTAL PERÍODO", _money(resumen["total_periodo"])),
    ]
    return _resumen_card("Resumen Facturación", filas)


def _jetsmart_resumen_html(resumen: dict) -> str:
    """Card con la hoja LIQUIDACION (y el total de CVLP) de JetSmart."""
    return _resumen_card("Liquidación JetSmart", [
        ("item", "Ventas totales", _money(resumen["ventas_totales"])),
        ("item", "IVA", _money(resumen["iva"])),
        ("subtotal", "Ventas Netas", _money(resumen["ventas_netas"])),
        ("item", "Comisiones por ventas — domésticas (7,5%)", _money(resumen["comision_domestica"])),
        ("item", "Comisiones por ventas — internacionales", _money(resumen["comision_inter"])),
        ("item", f"GHA Services ({_num(resumen['kg_total'])} kg)", _money(resumen["gha_services"])),
        ("item", "IVA de servicios y comisiones", _money(resumen["iva_servicios"])),
        ("item", "IIBB", _money(resumen["iibb"])),
        ("total", "Total a entregar a WCS", _money(resumen["total_wcs"])),
        ("item", "CVLP — Total final (con IVA 21% s/ neto gravado)", _money(resumen["cvlp"]["total_final"])),
    ])


def _jetsmart_avisos(resumen: dict) -> None:
    """Avisos amarillos de JetSmart: nunca bloquean el calculo."""
    activity = unconfirmed_tarifa_activity(resumen)
    if activity:
        # Agrupado por tarifa: hoy es una sola (0,185 en todas), no hace
        # falta repetirla estacion por estacion.
        por_tarifa: dict[float, list[str]] = {}
        for e in resumen["estaciones"]:
            if e["estacion"] in activity:
                por_tarifa.setdefault(e["tarifa_usd_kg"], []).append(e["estacion"])
        detalle_txt = "; ".join(
            f"{str(tarifa).replace('.', ',')} USD/kg en {', '.join(estaciones)}"
            for tarifa, estaciones in por_tarifa.items()
        )
        st.warning(
            f"Tarifa no confirmada: la tarifa GHA de JetSmart todavía no está confirmada con el cliente "
            f"(se aplicó {detalle_txt}). Revisá GHA Services antes de usar estos montos.",
            icon=":material/warning:",
        )


def _cruce_to_frames(cruce: dict) -> dict:
    """El cruce guardado en la base viene como listas de dicts; el recien
    calculado, como DataFrames. Normaliza a DataFrames."""
    return {
        "resumen": cruce["resumen"],
        "excepciones": pd.DataFrame(cruce["excepciones"]),
        "incluidas_ultimo_dia": pd.DataFrame(cruce["incluidas_ultimo_dia"]),
    }


def _render_cruce(cruce: dict, *, key: str) -> None:
    """Resultado del cruce con Ariel: resumen, aviso de la regla de ultimo
    dia y SOLO las excepciones (lo que matchea limpio no se muestra)."""
    r = cruce["resumen"]
    excepciones = cruce["excepciones"]
    ultimo_dia = cruce["incluidas_ultimo_dia"]
    motivos = r["motivos_inclusion"]

    n_exc = r["excepciones"]
    st.markdown(
        _resumen_card("Cruce con Ariel", [
            ("subtotal", "Guías que entran a la liquidación", _num(r["incluidas"], 0)),
            ("ok", "Matchean limpio (export = Ariel)", _num(motivos.get("match", 0), 0)),
            ("item", "Ya declaradas el mes anterior (quedan afuera)", _num(r["ya_declaradas_mes_anterior"], 0)),
            ("warn" if n_exc else "ok", "Excepciones para revisar" if n_exc else "Sin excepciones", _num(n_exc, 0)),
        ]),
        unsafe_allow_html=True,
    )

    if not ultimo_dia.empty:
        st.warning(
            f"Se incluyeron {len(ultimo_dia)} guías de Ariel del último día del mes ({r['ultimo_dia']}) "
            "por la regla de último día, sin confirmación cruzada: no están en el export de este mes y "
            "todavía no se cargó el export del mes siguiente para confirmarlas.",
            icon=":material/warning:",
        )
        with st.expander(f"Ver las {len(ultimo_dia)} guías incluidas por la regla de último día", key=f"{key}_ultimo_dia"):
            st.dataframe(ultimo_dia, use_container_width=True, hide_index=True, column_config={
                c: st.column_config.NumberColumn(format="localized") for c in ("Kg Ariel", "Ingreso Ariel")
            })

    if r["ingreso_fuera_por_revision"]:
        st.warning(
            f"Hay guías de Ariel a revisión manual que quedaron FUERA de la liquidación por "
            f"{_money(r['ingreso_fuera_por_revision'])} de ingreso. Revisalas en la tabla de excepciones.",
            icon=":material/warning:",
        )

    if excepciones.empty:
        st.caption("Sin excepciones: todas las guías cruzaron limpio.")
    else:
        st.markdown('<div class="hwc-group-title">Excepciones del cruce</div>', unsafe_allow_html=True)
        # Vacio en vez de "None": numeros como numero, texto en blanco.
        vista = excepciones.copy()
        for col in ["Kg export", "Kg Ariel", "Ingreso export", "Ingreso Ariel"]:
            vista[col] = pd.to_numeric(vista[col], errors="coerce")
        for col in ["Fecha Ariel", "Detalle"]:
            vista[col] = vista[col].fillna("")
        st.dataframe(vista, use_container_width=True, hide_index=True, key=f"{key}_excepciones", column_config={
            c: st.column_config.NumberColumn(format="localized")
            for c in ("Kg export", "Kg Ariel", "Ingreso export", "Ingreso Ariel")
        })


def _render_jetsmart_result(
    uploaded_file, tipo_cambio: float, vuelos_inter: dict[str, int], archivo_anterior, export_siguiente_file,
) -> tuple[object, tuple[str, str]]:
    """Corre el cruce con Ariel + la liquidacion de JetSmart y muestra el resultado.

    uploaded_file / archivo_anterior: archivos de trabajo de Anita (hoja BD
    primero + hoja ARIEL) del mes y del mes anterior. export_siguiente_file:
    export de guias del mes siguiente, opcional.
    """
    with st.spinner("Cruzando con Ariel y generando liquidación..."):
        export = load_jetsmart_export(uploaded_file)
        period, otros_meses = detect_jetsmart_period(export)
        ariel, _ = load_ariel(uploaded_file)
        export_anterior = load_jetsmart_export(archivo_anterior)
        ariel_anterior, _ = load_ariel(archivo_anterior)
        export_siguiente = load_jetsmart_export(export_siguiente_file) if export_siguiente_file is not None else None
        resultado = cruzar_jetsmart(export, ariel, period, export_anterior, ariel_anterior, export_siguiente)

        guias = build_jetsmart_guias(resultado.incluidas)
        resumen = build_jetsmart_resumen(guias, tipo_cambio, vuelos_inter)
        cruce = {
            "resumen": resultado.resumen,
            "excepciones": resultado.excepciones,
            "incluidas_ultimo_dia": resultado.incluidas_ultimo_dia,
        }

        buffer = io.BytesIO()
        write_jetsmart_liquidacion(guias, resumen, period, buffer, cruce=cruce)
        buffer.seek(0)

        guardar_liquidacion_jetsmart(
            nombre_archivo=uploaded_file.name,
            file_bytes=uploaded_file.getvalue(),
            period=period,
            filas_export=len(export),
            filas_otros_meses=len(otros_meses),
            guias=guias,
            resumen=resumen,
            parametros={"tipo_cambio": tipo_cambio, "vuelos_inter": vuelos_inter},
            cruce={
                "resumen": resultado.resumen,
                "excepciones": resultado.excepciones.astype(object).where(resultado.excepciones.notna(), None).to_dict(orient="records"),
                "incluidas_ultimo_dia": resultado.incluidas_ultimo_dia.to_dict(orient="records"),
            },
        )

    _render_exito("jetsmart", period, buffer)
    vuelos_txt = ", ".join(f"{station} {n}" for station, n in vuelos_inter.items())
    st.caption(
        f"Período detectado: {period_label(period)} · {len(guias)} guías · "
        f"TC {_num(tipo_cambio)} y vuelos internacionales ({vuelos_txt}) cargados a mano."
    )
    _jetsmart_avisos(resumen)
    _render_cruce(cruce, key="resultado_cruce")
    st.markdown(_jetsmart_resumen_html(resumen), unsafe_allow_html=True)
    return buffer, period


# ---------------------------------------------------------------------------
# Encabezado: marca + nombre de la herramienta, compacto, para que la accion
# principal (cargar un archivo) quede arriba, sin scrollear.
# ---------------------------------------------------------------------------
st.markdown(
    f"""
<div class="hwc-hero">
    <div class="hwc-hero-brand">
        <span class="hwc-hero-logo"><img src="{LOGO_DATA_URI}" alt="Handyway Cargo" /></span>
        <div>
            <div class="hwc-hero-title">Reportes HWC</div>
            <p class="hwc-hero-sub">Liquidaciones por aerolínea · Handyway Cargo</p>
        </div>
    </div>
</div>
""",
    unsafe_allow_html=True,
)


def _step(num: int, titulo: str, ayuda: str | None = None) -> None:
    ayuda_html = f'<div class="hwc-step-help">{ayuda}</div>' if ayuda else ""
    st.markdown(
        f'<div class="hwc-step"><span class="hwc-step-num">{num}</span><div>'
        f'<div class="hwc-step-title">{titulo}</div>{ayuda_html}</div></div>',
        unsafe_allow_html=True,
    )


def _checklist(requisitos: list[tuple[str, bool]]) -> None:
    """Que falta para poder procesar: explica por que el boton esta
    deshabilitado, en vez de dejarlo gris sin motivo."""
    if all(ok for _, ok in requisitos):
        st.markdown(
            '<div class="hwc-ready"><span class="hwc-dot hwc-dot-success">✓</span>'
            "Todo listo. Revisá que sea el archivo correcto y procesalo.</div>",
            unsafe_allow_html=True,
        )
        return
    items = "".join(
        f'<li class="{"hwc-req-ok" if ok else "hwc-req-pending"}">'
        f'<span class="hwc-dot {"hwc-dot-success" if ok else "hwc-dot-pending"}">{"✓" if ok else ""}</span>{texto}</li>'
        for texto, ok in requisitos
    )
    st.markdown(
        f'<div class="hwc-req"><div class="hwc-req-title">Para procesar falta completar:</div><ul>{items}</ul></div>',
        unsafe_allow_html=True,
    )


# ---------------------------------------------------------------------------
# Seccion 1: Cargar archivo nuevo -- formulario + resultado inmediato de
# procesarlo. Va junto y siempre visible porque es un solo flujo de
# principio a fin. Orden: aerolinea primero, porque el archivo que hay que
# subir depende de ella (JetSmart usa el archivo de trabajo, no el export).
# ---------------------------------------------------------------------------
with st.container(key="zona_carga"):
    st.markdown('<div class="hwc-section-title">Cargar archivo nuevo</div>', unsafe_allow_html=True)
    st.markdown(
        '<p class="hwc-section-sub">Elegí la aerolínea, subí el archivo y procesalo. El resultado aparece acá mismo.</p>',
        unsafe_allow_html=True,
    )

    with st.container(border=True, key="card_config"):
        _step(1, "Elegí la aerolínea")
        airline_key = st.segmented_control(
            "Aerolínea", options=sorted(AIRLINE_CONFIGS), format_func=_airline_name,
            key="aerolinea_sel", label_visibility="collapsed", width="stretch",
        )
    es_jetsmart_sel = airline_key is not None and AIRLINE_CONFIGS[airline_key].get("flow") == FLOW_JETSMART

    with st.container(border=True, key="card_upload"):
        _step(2, "Subí el archivo", _archivo_hint(airline_key))
        # Etiqueta fija + key: si cambia la aerolinea, el archivo ya subido
        # no se pierde (la ayuda dinamica va arriba, en _step).
        uploaded_file = st.file_uploader("Archivo del mes", type="xlsx", key="archivo_mes", label_visibility="collapsed")

        # JetSmart necesita ademas el archivo de trabajo del mes anterior
        # (cruce con Ariel) y, opcional, el export del mes siguiente.
        archivo_anterior = export_siguiente_file = None
        if es_jetsmart_sel:
            archivo_anterior = st.file_uploader(
                "Archivo de trabajo del mes anterior (LIQUIDACION ECS, con hojas BD y ARIEL)",
                type="xlsx", key="js_anterior",
            )
            export_siguiente_file = st.file_uploader(
                "Export de guías del mes siguiente (opcional: confirma las guías del último día del mes)",
                type="xlsx", key="js_siguiente",
            )

    # JetSmart necesita datos que no vienen en ningun export: el tipo de
    # cambio del periodo (todavia sin confirmar de donde sale -- por ahora
    # manual) y la cantidad de vuelos internacionales por estacion de los
    # manifiestos. Sin TC no se puede calcular GHA Services.
    tipo_cambio, vuelos_inter = 0.0, {}
    if es_jetsmart_sel:
        with st.container(border=True, key="card_datos"):
            _step(3, "Completá los datos del período", "No vienen en ningún archivo: salen del tipo de cambio del mes y de los manifiestos.")
            tipo_cambio = st.number_input("Tipo de cambio del período (ARS por USD)", min_value=0.0, value=0.0, step=1.0, format="%.2f")
            vuelos_cols = st.columns(len(JETSMART_COMISION_INTER_USD_POR_VUELO))
            for col, (station, usd) in zip(vuelos_cols, JETSMART_COMISION_INTER_USD_POR_VUELO.items()):
                with col:
                    vuelos_inter[station] = int(st.number_input(
                        f"Vuelos internacionales {station} ({usd:g} USD c/u)", min_value=0, value=0, step=1,
                    ))

    requisitos = [("Elegir la aerolínea", airline_key is not None), ("Subir el archivo del mes", uploaded_file is not None)]
    if es_jetsmart_sel:
        if uploaded_file is not None and not has_ariel_sheet(uploaded_file):
            requisitos[1] = ("El archivo del mes tiene que tener la hoja ARIEL (el que subiste no la tiene)", False)
        if archivo_anterior is not None and not has_ariel_sheet(archivo_anterior):
            requisitos.append(("El archivo del mes anterior tiene que tener la hoja ARIEL", False))
        else:
            requisitos.append(("Subir el archivo de trabajo del mes anterior", archivo_anterior is not None))
        requisitos.append(("Cargar el tipo de cambio", tipo_cambio > 0))
    listo = all(ok for _, ok in requisitos)

    with st.container(key="card_accion"):
        _checklist(requisitos)
        generate = st.button(
            f"Procesar y guardar · {_airline_name(airline_key)}" if airline_key else "Procesar y guardar",
            disabled=not listo, type="primary", icon=":material/play_arrow:", width="stretch",
        )

    # -----------------------------------------------------------------------
    # Resultado (misma seccion -- no hace falta ir a otro lado para ver lo
    # que se acaba de procesar). Si el archivo no es el esperado (ej. el
    # export de otra aerolinea), se explica en vez de mostrar un traceback.
    # -----------------------------------------------------------------------
    if generate:
        with st.container(border=True, key="card_result"):
            st.markdown('<div class="hwc-result-title">Resultado</div>', unsafe_allow_html=True)
            airline_cfg = AIRLINE_CONFIGS[airline_key]
            try:
                if airline_cfg.get("flow") == FLOW_JETSMART:
                    buffer, period = _render_jetsmart_result(
                        uploaded_file, tipo_cambio, vuelos_inter, archivo_anterior, export_siguiente_file,
                    )
                elif airline_cfg.get("flow") == FLOW_LIQUIDACION:
                    buffer, _, period = _render_liquidacion_result(uploaded_file)
                else:
                    with st.spinner("Generando reporte..."):
                        df = load_original(uploaded_file)
                        period, excluded = detect_period_exclusions(df, airline_key)
                        sheets = build_airline_report(df, airline_key, period=period)

                        buffer = io.BytesIO()
                        write_report(sheets, buffer)
                        buffer.seek(0)

                        guardar_reporte_simple(
                            nombre_archivo=uploaded_file.name,
                            file_bytes=uploaded_file.getvalue(),
                            df=df,
                            airline_key=airline_key,
                            period=period,
                            sheets=sheets,
                        )

                    _render_exito(airline_key, period, buffer)
                    st.caption(f"Período detectado: {period_label(period)}")

                    if not excluded.empty:
                        breakdown = excluded[COL_COD_VUELO].map(extract_period).value_counts()
                        detalle_txt = ", ".join(f"{n} de {period_label(p)}" for p, n in breakdown.items())
                        fila_word = "fila" if len(excluded) == 1 else "filas"
                        st.warning(
                            f"Se excluyeron {len(excluded)} {fila_word} fuera del período detectado "
                            f"({period_label(period)}) del reporte de {airline_key.upper()}: {detalle_txt}.",
                            icon=":material/warning:",
                        )

                    unconfirmed_stations = airline_cfg.get("unconfirmed_stations", [])
                    if unconfirmed_stations:
                        activity = detect_unconfirmed_activity(sheets, unconfirmed_stations)
                        if activity:
                            estaciones = ", ".join(activity)
                            st.warning(
                                f"Se detectaron movimientos en {estaciones} para {airline_key.upper()}. "
                                f"La tarifa aplicada ahí todavía no está confirmada con el cliente — "
                                f"revisá los montos manualmente antes de usarlos.",
                                icon=":material/warning:",
                            )

                    for charge_type_key in airline_cfg["charge_types"]:
                        _, label = CHARGE_TYPE_LABELS.get(charge_type_key, ("", charge_type_key))
                        group_sheets = _sheets_for_charge_type(sheets, charge_type_key)
                        st.markdown(
                            _resumen_card(label, [
                                ("vacio", f"{sheet_name} · sin movimiento", None) if len(sheet_df) == 0
                                else ("ok", sheet_name, f"{_num(len(sheet_df), 0)} filas")
                                for sheet_name, sheet_df in group_sheets.items()
                            ]),
                            unsafe_allow_html=True,
                        )
            except Exception as exc:  # el detalle tecnico va aparte, chico
                st.error(
                    f"**No se pudo procesar el archivo como {_airline_name(airline_key)}.** "
                    "Revisá que sea el archivo correcto para esa aerolínea y volvé a intentarlo.",
                    icon=":material/error:",
                )
                st.caption(f"Detalle técnico: {type(exc).__name__}: {exc}")


# ---------------------------------------------------------------------------
# Seccion 2: Historial de Liquidaciones -- vista de SOLO LECTURA sobre
# liquidaciones ya persistidas (ver obtener_filtros_historial/
# obtener_historial en src/db.py). No recalcula nada -- muestra tal cual lo
# que guardo cada "Procesar y guardar" pasado. Arranca colapsada: quien
# acaba de procesar un archivo no tiene por que encontrarse con todo el
# dashboard historico; se despliega solo si se quiere consultar algo.
# ---------------------------------------------------------------------------
with st.container(key="zona_historial"):
    st.markdown('<div class="hwc-section-title">Historial de Liquidaciones</div>', unsafe_allow_html=True)
    st.markdown(
        '<p class="hwc-section-sub">Consulta de solo lectura sobre lo ya generado y guardado — no vuelve a calcular nada.</p>',
        unsafe_allow_html=True,
    )

    # La etiqueta del desplegable resume que hay adentro, para que invite a
    # abrirlo. Misma consulta fail-soft de siempre (None = base caida).
    historial_todo = obtener_historial()
    if historial_todo is None:
        etiqueta_historial = "Ver historial — no disponible ahora (sin conexión a la base de datos)"
    elif historial_todo.empty:
        etiqueta_historial = "Ver historial — todavía no hay liquidaciones guardadas"
    else:
        n_vig = int(historial_todo["es_vigente"].sum())
        ultima = historial_todo["generado_en"].max().strftime("%d/%m/%Y")
        etiqueta_historial = (
            f"Ver historial — {n_vig} liquidaci{'ón vigente' if n_vig == 1 else 'ones vigentes'}"
            f" · última el {ultima}"
        )

    with st.expander(etiqueta_historial, expanded=False, key="expander_historial", icon=":material/history:"):
        filtros = obtener_filtros_historial()

        if filtros is None:
            st.info(
                "El historial no está disponible en este momento (sin conexión a la base de datos). "
                "Podés seguir generando y descargando reportes con normalidad desde \"Cargar archivo nuevo\".",
                icon=":material/cloud_off:",
            )
        elif not filtros["aerolineas"]:
            st.info("Todavía no hay ningún reporte generado guardado en el historial.", icon=":material/inbox:")
        else:
            col_aerolinea, col_periodo, col_estacion, col_anteriores = st.columns([1, 1, 1, 0.9], vertical_alignment="bottom")
            with col_aerolinea:
                aerolinea_opciones = {"Todas": None} | {_airline_name(a): a for a in filtros["aerolineas"]}
                aerolinea_sel = st.selectbox("Aerolínea", list(aerolinea_opciones), key="hist_aerolinea")
            with col_periodo:
                periodo_opciones = {"Todos": None}
                for periodo_fecha, mes, anio in filtros["periodos"]:
                    periodo_opciones[period_label((mes, anio))] = periodo_fecha
                periodo_sel_label = st.selectbox("Período", list(periodo_opciones.keys()), key="hist_periodo")
            with col_estacion:
                estacion_opciones = ["Todas"] + filtros["estaciones"]
                estacion_sel = st.selectbox("Estación", estacion_opciones, key="hist_estacion")
            with col_anteriores:
                ver_anteriores = st.toggle("Mostrar generaciones anteriores", value=False, key="hist_anteriores")

            aerolinea_filtro = aerolinea_opciones[aerolinea_sel]
            periodo_filtro = periodo_opciones[periodo_sel_label]
            estacion_filtro = None if estacion_sel == "Todas" else estacion_sel

            historial_df = obtener_historial(aerolinea_filtro, periodo_filtro, estacion_filtro)

            if historial_df is None:
                st.warning("No se pudo consultar el historial ahora mismo. Probá de nuevo en unos minutos.", icon=":material/cloud_off:")
            elif historial_df.empty:
                st.info("No hay liquidaciones para los filtros seleccionados.", icon=":material/search_off:")
            else:
                # El total SOLO suma la generacion mas reciente de cada
                # combinacion (aerolinea/estacion/tipo_cargo/periodo) --
                # es_vigente lo calcula obtener_historial() con ROW_NUMBER().
                # Si algo se regenero, las anteriores se pueden ver en la
                # tabla (toggle) para trazabilidad, pero nunca suman: el
                # numero grande no debe duplicar plata.
                vigentes = historial_df[historial_df["es_vigente"]]
                total = float(vigentes["monto_total"].fillna(0).sum())
                cantidad_vigente = len(vigentes)
                cantidad_anteriores = len(historial_df) - cantidad_vigente
                ultima_fecha = historial_df["generado_en"].max().strftime("%d/%m/%Y %H:%M")

                st.markdown(
                    '<div class="hwc-kpis">'
                    f'<div class="hwc-kpi hwc-kpi-hero"><div class="hwc-kpi-label">Total vigente</div>'
                    f'<div class="hwc-kpi-value">{_money(total)}</div>'
                    f'<div class="hwc-kpi-note">Suma de la última generación de cada liquidación</div></div>'
                    f'<div class="hwc-kpi"><div class="hwc-kpi-label">Liquidaciones vigentes</div>'
                    f'<div class="hwc-kpi-value">{cantidad_vigente}</div>'
                    f'<div class="hwc-kpi-note">{cantidad_anteriores} generaci{"ón anterior" if cantidad_anteriores == 1 else "ones anteriores"} (no suman)</div></div>'
                    f'<div class="hwc-kpi"><div class="hwc-kpi-label">Aerolíneas</div>'
                    f'<div class="hwc-kpi-value">{vigentes["aerolinea"].nunique()}</div>'
                    f'<div class="hwc-kpi-note">con liquidaciones en este filtro</div></div>'
                    f'<div class="hwc-kpi"><div class="hwc-kpi-label">Último procesamiento</div>'
                    f'<div class="hwc-kpi-value hwc-kpi-value-sm">{ultima_fecha}</div>'
                    f'<div class="hwc-kpi-note">hora UTC</div></div>'
                    '</div>',
                    unsafe_allow_html=True,
                )

                # ---------------------------------------------------------
                # Dos graficos, ambos SOLO con montos vigentes (misma
                # logica que el total de arriba) para no duplicar plata.
                # ---------------------------------------------------------
                col_chart1, col_chart2 = st.columns(2, gap="large")
                with col_chart1:
                    with st.container(key="card_chart_aerolinea"):
                        st.markdown('<div class="hwc-chart-title">Total facturado por aerolínea</div>', unsafe_allow_html=True)
                        por_aerolinea = (
                            vigentes.assign(_aerolinea=vigentes["aerolinea"].map(_airline_name))
                            .groupby("_aerolinea")["monto_total"]
                            .sum()
                        )
                        st.altair_chart(_grafico_barras(por_aerolinea, "Aerolínea", horizontal=True), use_container_width=True)

                with col_chart2:
                    with st.container(key="card_chart_mes"):
                        st.markdown('<div class="hwc-chart-title">Evolución por mes</div>', unsafe_allow_html=True)
                        # A proposito ignora el filtro de periodo (pasa None):
                        # el punto de este grafico es mostrar varios meses a la
                        # vez, asi que no tiene sentido dejar que el propio
                        # filtro de periodo lo colapse a una sola barra. Si
                        # respeta aerolinea/estacion, igual que el de al lado.
                        evolucion_df = obtener_historial(aerolinea_filtro, None, estacion_filtro)
                        if evolucion_df is None:
                            st.caption("No se pudo cargar la evolución mensual ahora mismo.")
                        else:
                            evolucion_vigente = evolucion_df[evolucion_df["es_vigente"]].copy()
                            evolucion_vigente["_periodo_label"] = [
                                period_label((mes, anio))
                                for mes, anio in zip(evolucion_vigente["periodo_mes"], evolucion_vigente["periodo_anio"])
                            ]
                            por_mes = (
                                evolucion_vigente.sort_values("periodo")
                                .groupby("_periodo_label", sort=False)["monto_total"]
                                .sum()
                            )
                            # Orden cronologico (no alfabetico -- "julio"
                            # quedaria antes que "mayo"): se pasa la lista ya
                            # ordenada como "sort" del eje de Altair.
                            st.altair_chart(
                                _grafico_barras(por_mes, "Período", horizontal=False, orden=list(por_mes.index)),
                                use_container_width=True,
                            )
                            if len(por_mes) == 1:
                                st.caption("Todavía hay un solo período cargado — el gráfico va a sumar meses a medida que se generen más reportes.")

                # Tabla: por defecto solo lo vigente (lo que suma); las
                # generaciones anteriores quedan a un toggle de distancia.
                vista_df = (historial_df if ver_anteriores else vigentes).reset_index(drop=True)
                st.markdown(
                    f'<div class="hwc-chart-title" style="margin-top:0.4rem;">Liquidaciones '
                    f'<span class="hwc-chart-sub">{len(vista_df)} '
                    f'{"en total, incluidas las anteriores" if ver_anteriores else "vigentes"}</span></div>',
                    unsafe_allow_html=True,
                )
                tabla = pd.DataFrame({
                    "Estado": vista_df["es_vigente"].map(lambda v: "✓ Vigente" if v else "Anterior"),
                    "Generada (UTC)": vista_df["generado_en"].dt.strftime("%d/%m/%Y %H:%M"),
                    "Aerolínea": vista_df["aerolinea"].map(_airline_name),
                    "Estación": vista_df["estacion"],
                    "Tipo de cargo": vista_df["tipo_cargo"].map(lambda t: CHARGE_TYPE_LABELS.get(t, ("", t))[1]),
                    "Período": [
                        period_label((mes, anio)) for mes, anio in zip(vista_df["periodo_mes"], vista_df["periodo_anio"])
                    ],
                    "Movimientos": vista_df["cantidad_filas"],
                    "Monto total": vista_df["monto_total"].fillna(0),
                })
                st.dataframe(
                    tabla, use_container_width=True, hide_index=True,
                    height=min(420, 35 * (len(tabla) + 1) + 3),
                    column_config={
                        "Movimientos": st.column_config.NumberColumn(format="localized"),
                        "Monto total": st.column_config.NumberColumn(format="dollar"),
                    },
                )

                # ---------------------------------------------------------
                # Detalle linea por linea de UNA liquidacion puntual. El
                # combo referencia filas de vista_df (lo que se ve en la
                # tabla) por posicion -- alcanza porque se reconstruye en
                # cada rerun a partir del mismo query, en el mismo orden.
                # ---------------------------------------------------------
                with st.container(border=True, key="card_detalle"):
                    st.markdown(
                        '<div class="hwc-chart-title">Detalle de una liquidación</div>'
                        '<p class="hwc-detail-help">Elegí una de la tabla para ver su detalle línea por línea y descargar su Excel.</p>',
                        unsafe_allow_html=True,
                    )

                    def _etiqueta_detalle(i: int) -> str:
                        fila = vista_df.iloc[i]
                        tipo_label = CHARGE_TYPE_LABELS.get(fila["tipo_cargo"], ("", fila["tipo_cargo"]))[1]
                        periodo_txt = period_label((fila["periodo_mes"], fila["periodo_anio"]))
                        fecha_txt = fila["generado_en"].strftime("%d/%m/%Y %H:%M UTC")
                        vigencia_txt = "" if fila["es_vigente"] else " (anterior)"
                        return (
                            f"{_airline_name(fila['aerolinea'])} · {fila['estacion']} · {tipo_label} · "
                            f"{periodo_txt} · {fecha_txt}{vigencia_txt}"
                        )

                    seleccion = st.selectbox(
                        "Elegí una liquidación para ver su detalle línea por línea",
                        options=range(len(vista_df)),
                        format_func=_etiqueta_detalle,
                        label_visibility="collapsed",
                        key="hist_detalle_selector",
                    )
                    fila_sel = vista_df.iloc[seleccion]

                    if fila_sel["cantidad_filas"] == 0:
                        st.info("Esta liquidación no tiene movimientos asociados (sin movimiento).", icon=":material/info:")
                    else:
                        with st.spinner("Cargando detalle..."):
                            resultado_detalle = _obtener_detalle_liquidacion(fila_sel)
                        if resultado_detalle is None:
                            st.warning(
                                "No se pudo cargar el detalle ahora mismo. Probá de nuevo en unos minutos.",
                                icon=":material/warning:",
                            )
                        else:
                            detalle_df, resumen, sheet_name = resultado_detalle
                            es_jetsmart = fila_sel["tipo_cargo"] == "jetsmart_liquidacion"

                            cruce_guardado = _cruce_to_frames(resumen["_cruce"]) if es_jetsmart and resumen.get("_cruce") else None
                            if es_jetsmart:
                                if cruce_guardado is not None:
                                    _render_cruce(cruce_guardado, key="hist_cruce")
                                st.markdown(_jetsmart_resumen_html(resumen), unsafe_allow_html=True)
                                st.caption(
                                    f"{len(detalle_df)} guías (hoja \"GUIAS\") — TC {_num(resumen['tipo_cambio'])}. "
                                    "Mismo cálculo que el Excel real de esta liquidación (jetsmart_builder.py)."
                                )
                            elif resumen is not None:
                                # Resumen Facturacion (solo LATAM): mismo
                                # desglose de IVA que trae la hoja real del
                                # Excel -- sub-items de cada sub-factura,
                                # TOTAL LA, IVA, TOTAL 4M y TOTAL PERIODO.
                                # build_latam_resumen() es la MISMA funcion que
                                # usa write_liquidacion() para el Excel real
                                # (ver _write_resumen_sheet en
                                # liquidacion_builder.py), no una
                                # reimplementacion del calculo de IVA.
                                st.markdown(_latam_resumen_html(resumen), unsafe_allow_html=True)
                                st.caption(
                                    f"{len(detalle_df)} filas de detalle (hoja \"Detalle de Facturación\") — "
                                    "mismo filtrado y mismo cálculo de IVA que el Excel real de esta liquidación "
                                    "(liquidacion_builder.py), sin recalcular nada distinto."
                                )
                            else:
                                st.caption(
                                    f"{len(detalle_df)} filas — mismo filtrado que arma el Excel real de esta "
                                    "liquidación (report_builder.py), sin recalcular nada."
                                )

                            # Excel en memoria con las mismas funciones que arman el
                            # Excel real (write_liquidacion / write_report) -- no
                            # reimplementa el armado del archivo, solo lo reusa
                            # sobre el mismo detalle ya reconstruido arriba.
                            detalle_buffer = io.BytesIO()
                            if es_jetsmart:
                                write_jetsmart_liquidacion(
                                    detalle_df, resumen, (fila_sel["periodo_mes"], fila_sel["periodo_anio"]), detalle_buffer,
                                    cruce=cruce_guardado,
                                )
                            elif resumen is not None:
                                write_liquidacion(detalle_df, resumen, detalle_buffer)
                            else:
                                write_report({sheet_name: detalle_df}, detalle_buffer)
                            detalle_buffer.seek(0)

                            periodo_slug = period_slug((fila_sel["periodo_mes"], fila_sel["periodo_anio"]))
                            st.download_button(
                                label="Descargar Excel de esta liquidación",
                                icon=":material/download:",
                                data=detalle_buffer,
                                file_name=f"{fila_sel['aerolinea']}_{fila_sel['estacion']}_{periodo_slug}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                type="secondary",
                                on_click="ignore",
                                key="hist_detalle_download",
                            )

                            # Alto dinamico hasta un tope: con liquidaciones
                            # chicas (ej. 3 filas) no deja un montón de grilla
                            # vacia; con liquidaciones grandes (ej. 1033 filas)
                            # se topea en 420px y scrollea adentro del recuadro
                            # en vez de estirar la pagina entera.
                            alto_tabla = min(420, 38 * (len(detalle_df) + 1) + 4)
                            st.dataframe(
                                detalle_df, use_container_width=True, hide_index=True, height=alto_tabla,
                                column_config={
                                    col: st.column_config.NumberColumn(format="localized")
                                    for col in detalle_df.columns if pd.api.types.is_float_dtype(detalle_df[col])
                                },
                            )


st.markdown('<div class="hwc-footer">Handyway Cargo · Automatización de reportes</div>', unsafe_allow_html=True)
