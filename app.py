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

h1, h2, h3 {{ font-family: 'Inter', sans-serif; font-weight: 800; color: var(--hwc-blue-text); letter-spacing: -0.01em; }}
/* Nota: no pisar el color de todos los <p> de stMarkdownContainer aca:
   Streamlit tambien usa ese mismo contenedor para el texto de los botones,
   asi que una regla global de color rompe el texto blanco de los botones
   y del banner. El color de cuerpo normal ya se hereda de html/body. */
[data-testid="stCaptionContainer"] {{ color: var(--hwc-text-muted) !important; }}
[data-testid="stWidgetLabel"] p {{ font-weight: 600; color: var(--hwc-text); font-size: 0.85rem; }}

/* --- Titulo principal, con el logo al lado del nombre --- */
.hwc-page-title {{
    display: flex; align-items: center; gap: 0.85rem;
    font-size: 2.4rem; font-weight: 800; color: var(--hwc-blue-text);
    letter-spacing: -0.02em; line-height: 1.1;
    margin: 0 0 1.2rem 0;
}}
.hwc-page-title-logo {{ height: 3rem; width: auto; flex-shrink: 0; }}

/* --- Banner superior de marca --- */
.hwc-hero {{
    background: linear-gradient(135deg, var(--hwc-blue-deep) 0%, var(--hwc-blue) 100%);
    border-radius: 16px;
    padding: 1.8rem 2.1rem;
    margin-bottom: 1.6rem;
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
.hwc-hero-sub {{ color: rgba(255, 255, 255, 0.82) !important; font-size: 0.85rem; margin: 0; position: relative; }}
.hwc-hero-tag {{
    display: inline-block; margin-top: 0.9rem;
    background: rgba(255, 255, 255, 0.14); border: 1px solid rgba(255, 255, 255, 0.35);
    color: #fff; font-size: 0.7rem; font-weight: 700; letter-spacing: 0.06em;
    text-transform: uppercase; padding: 0.28rem 0.75rem; border-radius: 100px; position: relative;
}}

/* --- Cards de seccion (subir archivo / aerolinea / resultado / login / historial) --- */
.st-key-card_upload[data-testid="stVerticalBlock"],
.st-key-card_config[data-testid="stVerticalBlock"],
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
.hwc-section-sub {{ color: var(--hwc-text-muted); font-size: 0.85rem; margin: 0 0 1rem 0.95rem; }}

.hwc-step {{
    display: flex; align-items: center; gap: 0.65rem;
    font-weight: 700; font-size: 1.05rem; color: var(--hwc-blue-text);
    margin-bottom: 1rem;
}}
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
.hwc-dot-empty {{ background: #E4E9F0; color: #9AA5B1; }}
.hwc-row-active {{ color: var(--hwc-text); }}
.hwc-row-empty {{ color: var(--hwc-text-muted); }}
.hwc-count {{
    margin-left: auto; font-weight: 700; color: var(--hwc-blue-text);
    background: #EAF1F6; padding: 0.15rem 0.65rem; border-radius: 100px; font-size: 0.78rem;
}}

/* --- Pantalla de acceso --- */
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

/* --- Footer --- */
.hwc-footer {{
    text-align: center; color: var(--hwc-text-muted); font-size: 0.78rem;
    margin-top: 1.6rem; padding-top: 1rem; border-top: 1px solid var(--hwc-border);
}}

/* --- Responsive: pantallas chicas --- */
@media (max-width: 480px) {{
    [data-testid="stMainBlockContainer"] {{ padding-left: 1rem !important; padding-right: 1rem !important; }}
    .hwc-page-title {{ font-size: 1.7rem; gap: 0.55rem; }}
    .hwc-page-title-logo {{ height: 2.2rem; }}
    .hwc-hero {{ padding: 1.4rem 1.3rem; }}
    .st-key-card_upload[data-testid="stVerticalBlock"],
    .st-key-card_config[data-testid="stVerticalBlock"],
    .st-key-card_result[data-testid="stVerticalBlock"],
    .st-key-card_login[data-testid="stVerticalBlock"],
    .st-key-card_historial[data-testid="stVerticalBlock"] {{ padding: 1.2rem 1.15rem; }}
    .hwc-watermark {{ background-size: 320px; background-position: bottom -30px right -30px; }}
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
                st.error("Contraseña incorrecta.", icon="🚫")
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


def _money(value: float) -> str:
    return f"$ {value:,.2f}"


def _grafico_barras(serie: pd.Series, nombre_categoria: str, color: str, orden: list[str] | None = None) -> alt.Chart:
    """Grafico de barras del mini-dashboard del Historial, en Altair
    directo (no st.bar_chart) para poder controlar tres cosas que
    st.bar_chart no expone via parametros:
    - Eje Y arrancando siempre en cero (scale zero=True explicito).
    - Tooltip con etiquetas en español y formato de moneda, en vez de los
      nombres de columna crudos (ej. "monto_total").
    - Orden explicito del eje X cuando corresponde (ver "orden": la
      evolucion por mes necesita orden cronologico, no alfabetico).
    """
    df = serie.reset_index()
    df.columns = [nombre_categoria, "Total facturado"]
    x_encoding = alt.X(
        f"{nombre_categoria}:N",
        title=None,
        sort=orden if orden is not None else "-y",
    )
    return (
        alt.Chart(df)
        .mark_bar(color=color)
        .encode(
            x=x_encoding,
            y=alt.Y("Total facturado:Q", title=None, scale=alt.Scale(zero=True)),
            tooltip=[
                alt.Tooltip(f"{nombre_categoria}:N", title=nombre_categoria),
                alt.Tooltip("Total facturado:Q", title="Total facturado", format="$,.2f"),
            ],
        )
    )


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

    st.success("Liquidación generada correctamente.", icon="✅")
    st.success("Guardado — ya lo podés consultar abajo, en Historial de Liquidaciones.", icon="✅")
    st.caption(f"Período detectado: {period_label(period)}")

    if not excluded.empty:
        breakdown = excluded[COL_COD_VUELO].map(extract_period).value_counts()
        detalle_txt = ", ".join(f"{n} de {period_label(p)}" for p, n in breakdown.items())
        fila_word = "fila" if len(excluded) == 1 else "filas"
        st.warning(
            f"Se excluyeron {len(excluded)} {fila_word} fuera del período detectado "
            f"({period_label(period)}) de LATAM: {detalle_txt}.",
            icon="⚠️",
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
                icon="⚠️",
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

    return buffer, resumen, period

def _jetsmart_resumen_html(resumen: dict) -> str:
    """Card con la hoja LIQUIDACION (y el total de CVLP) de JetSmart."""
    def row(label, value, *, bold=False, dot="✓", style=""):
        label_html = f"<b>{label}</b>" if bold else label
        return (
            f'<div class="hwc-row hwc-row-active"{style}><span class="hwc-dot hwc-dot-active">{dot}</span>'
            f'{label_html}<span class="hwc-count">{_money(value)}</span></div>'
        )

    separador = ' style="margin-top:0.5rem;border-top:1px solid var(--hwc-border);padding-top:0.6rem;"'
    return (
        '<div class="hwc-group-card">'
        '<div class="hwc-group-title">🧾 Liquidación JetSmart</div>'
        + row("Ventas totales", resumen["ventas_totales"])
        + row("IVA", resumen["iva"], dot="%")
        + row("Ventas Netas", resumen["ventas_netas"], bold=True, dot="Σ")
        + row("Comisiones por ventas — domésticas (7,5%)", resumen["comision_domestica"])
        + row("Comisiones por ventas — internacionales", resumen["comision_inter"])
        + row(f"GHA Services ({resumen['kg_total']:,.2f} kg)", resumen["gha_services"])
        + row("IVA de servicios y comisiones", resumen["iva_servicios"])
        + row("IIBB", resumen["iibb"])
        + row("Total a entregar a WCS", resumen["total_wcs"], bold=True, style=separador)
        + row("CVLP — Total final (con IVA 21% s/ neto gravado)", resumen["cvlp"]["total_final"])
        + "</div>"
    )


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
            icon="⚠️",
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

    filas = [
        ("Guías que entran a la liquidación", f'{r["incluidas"]}'),
        ("Matchean limpio (export = Ariel)", f'{motivos.get("match", 0)}'),
        ("Ya declaradas el mes anterior (quedan afuera)", f'{r["ya_declaradas_mes_anterior"]}'),
        ("Excepciones para revisar", f'{r["excepciones"]}'),
    ]
    rows_html = "".join(
        f'<div class="hwc-row hwc-row-active"><span class="hwc-dot hwc-dot-active">✓</span>'
        f'{label}<span class="hwc-count">{valor}</span></div>'
        for label, valor in filas
    )
    st.markdown(
        f'<div class="hwc-group-card"><div class="hwc-group-title">🔀 Cruce con Ariel</div>{rows_html}</div>',
        unsafe_allow_html=True,
    )

    if not ultimo_dia.empty:
        st.warning(
            f"Se incluyeron {len(ultimo_dia)} guías de Ariel del último día del mes ({r['ultimo_dia']}) "
            "por la regla de último día, sin confirmación cruzada: no están en el export de este mes y "
            "todavía no se cargó el export del mes siguiente para confirmarlas.",
            icon="⚠️",
        )
        with st.expander(f"Ver las {len(ultimo_dia)} guías incluidas por la regla de último día", key=f"{key}_ultimo_dia"):
            st.dataframe(ultimo_dia, use_container_width=True, hide_index=True)

    if r["ingreso_fuera_por_revision"]:
        st.warning(
            f"Hay guías de Ariel a revisión manual que quedaron FUERA de la liquidación por "
            f"{_money(r['ingreso_fuera_por_revision'])} de ingreso. Revisalas en la tabla de excepciones.",
            icon="⚠️",
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
        st.dataframe(vista, use_container_width=True, hide_index=True, key=f"{key}_excepciones")


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

    st.success("Liquidación generada correctamente.", icon="✅")
    st.success("Guardado — ya lo podés consultar abajo, en Historial de Liquidaciones.", icon="✅")
    vuelos_txt = ", ".join(f"{station} {n}" for station, n in vuelos_inter.items())
    st.caption(
        f"Período detectado: {period_label(period)} · {len(guias)} guías · "
        f"TC {tipo_cambio:,.2f} y vuelos internacionales ({vuelos_txt}) cargados a mano."
    )
    _jetsmart_avisos(resumen)
    _render_cruce(cruce, key="resultado_cruce")
    st.markdown(_jetsmart_resumen_html(resumen), unsafe_allow_html=True)
    return buffer, period


# ---------------------------------------------------------------------------
# Titulo principal (jerarquia por encima del banner de marca, sin tocarlo)
# ---------------------------------------------------------------------------
st.markdown(
    f'<div class="hwc-page-title"><img class="hwc-page-title-logo" src="{LOGO_DATA_URI}" alt="Handyway Cargo" />Handyway Cargo</div>',
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# Banner de marca
# ---------------------------------------------------------------------------
st.markdown(
    """
<div class="hwc-hero">
    <p class="hwc-hero-sub">Cargá un archivo nuevo para procesar, o consultá el historial de liquidaciones ya generadas.</p>
    <span class="hwc-hero-tag">Liquidaciones Handyway Cargo</span>
</div>
""",
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# Seccion 1: Cargar archivo nuevo -- formulario + resultado inmediato de
# procesarlo. Va junto y siempre visible porque es un solo flujo de
# principio a fin.
# ---------------------------------------------------------------------------
with st.container(key="zona_carga"):
    st.markdown('<div class="hwc-section-title">📥 Cargar archivo nuevo</div>', unsafe_allow_html=True)
    st.markdown(
        '<p class="hwc-section-sub">Subí el export, elegí la aerolínea y procesalo. El resultado aparece acá mismo.</p>',
        unsafe_allow_html=True,
    )

    with st.container(border=True, key="card_upload"):
        st.markdown('<div class="hwc-step"><span class="hwc-step-num">1</span>📄 Subí el archivo</div>', unsafe_allow_html=True)
        uploaded_file = st.file_uploader("Archivo original (export del sistema)", type="xlsx", label_visibility="collapsed")

    # -----------------------------------------------------------------------
    # Seccion 2: elegir aerolinea + generar
    # -----------------------------------------------------------------------
    with st.container(border=True, key="card_config"):
        st.markdown('<div class="hwc-step"><span class="hwc-step-num">2</span>✈️ Elegí la aerolínea</div>', unsafe_allow_html=True)
        airline_key = st.selectbox("Aerolínea", options=sorted(AIRLINE_CONFIGS), format_func=str.upper, label_visibility="collapsed")

        # JetSmart necesita datos que no vienen en ningun export: el tipo
        # de cambio del periodo (todavia sin confirmar de donde sale -- por
        # ahora manual) y la cantidad de vuelos internacionales por estacion
        # de los manifiestos. Sin TC no se puede calcular GHA Services.
        faltan_datos = False
        if AIRLINE_CONFIGS[airline_key].get("flow") == FLOW_JETSMART:
            st.caption(
                "Para JetSmart, el archivo del paso 1 es el archivo de trabajo del mes (LIQUIDACION ECS): "
                "la hoja BD con el export de guías primero, y la hoja ARIEL con el archivo de Ariel."
            )
            archivo_anterior = st.file_uploader(
                "Archivo de trabajo del mes anterior (LIQUIDACION ECS, con hojas BD y ARIEL) — obligatorio",
                type="xlsx", key="js_anterior",
            )
            export_siguiente_file = st.file_uploader(
                "Export de guías del mes siguiente — opcional, confirma las guías del último día del mes",
                type="xlsx", key="js_siguiente",
            )
            tipo_cambio = st.number_input("Tipo de cambio del período (ARS por USD)", min_value=0.0, value=0.0, step=1.0, format="%.2f")
            vuelos_cols = st.columns(len(JETSMART_COMISION_INTER_USD_POR_VUELO))
            vuelos_inter = {}
            for col, (station, usd) in zip(vuelos_cols, JETSMART_COMISION_INTER_USD_POR_VUELO.items()):
                with col:
                    vuelos_inter[station] = int(st.number_input(
                        f"Vuelos internacionales {station} ({usd:g} USD c/u)", min_value=0, value=0, step=1,
                    ))
            faltantes = []
            if uploaded_file is not None and not has_ariel_sheet(uploaded_file):
                faltantes.append("un archivo del mes con la hoja ARIEL (el del paso 1 no la tiene)")
            if archivo_anterior is None:
                faltantes.append("el archivo de trabajo del mes anterior")
            elif not has_ariel_sheet(archivo_anterior):
                faltantes.append("un archivo del mes anterior con la hoja ARIEL")
            if tipo_cambio <= 0:
                faltantes.append("el tipo de cambio")
            faltan_datos = bool(faltantes)
            if faltan_datos:
                st.caption("Para poder procesar falta: " + "; ".join(faltantes) + ".")

        generate = st.button("Procesar y guardar", disabled=uploaded_file is None or faltan_datos, type="primary")

    # -----------------------------------------------------------------------
    # Seccion 3: resultado (misma seccion -- no hace falta ir a otro lado
    # para ver lo que se acaba de procesar).
    # -----------------------------------------------------------------------
    if generate:
        with st.container(border=True, key="card_result"):
            st.markdown('<div class="hwc-step"><span class="hwc-step-num">3</span>📊 Resultado</div>', unsafe_allow_html=True)

            airline_cfg = AIRLINE_CONFIGS[airline_key]

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

                st.success("Reporte generado correctamente.", icon="✅")
                st.success("Guardado — ya lo podés consultar abajo, en Historial de Liquidaciones.", icon="✅")
                st.caption(f"Período detectado: {period_label(period)}")

                if not excluded.empty:
                    breakdown = excluded[COL_COD_VUELO].map(extract_period).value_counts()
                    detalle_txt = ", ".join(f"{n} de {period_label(p)}" for p, n in breakdown.items())
                    fila_word = "fila" if len(excluded) == 1 else "filas"
                    st.warning(
                        f"Se excluyeron {len(excluded)} {fila_word} fuera del período detectado "
                        f"({period_label(period)}) del reporte de {airline_key.upper()}: {detalle_txt}.",
                        icon="⚠️",
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
                            icon="⚠️",
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

            st.caption("Descarga opcional — la liquidación ya quedó guardada.")
            st.download_button(
                label="⬇️ Descargar Excel",
                data=buffer,
                file_name=f"{airline_key}_{period_slug(period)}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="secondary",
            )


# ---------------------------------------------------------------------------
# Seccion 2: Historial de Liquidaciones -- vista de SOLO LECTURA sobre
# liquidaciones ya persistidas (ver obtener_filtros_historial/
# obtener_historial en src/db.py). No recalcula nada -- muestra tal cual lo
# que guardo cada "Procesar y guardar" pasado. Arranca colapsada: quien
# acaba de procesar un archivo no tiene por que encontrarse con todo el
# dashboard historico; se despliega solo si se quiere consultar algo.
# ---------------------------------------------------------------------------
with st.container(key="zona_historial"):
    st.markdown('<div class="hwc-section-title">📊 Historial de Liquidaciones</div>', unsafe_allow_html=True)
    st.markdown(
        '<p class="hwc-section-sub">Consulta de solo lectura sobre lo ya generado y guardado — no vuelve a calcular nada.</p>',
        unsafe_allow_html=True,
    )

    with st.expander("Ver historial (filtros, gráficos y detalle de cada liquidación)", expanded=False, key="expander_historial"):
        with st.container(border=True, key="card_historial"):

            filtros = obtener_filtros_historial()

            if filtros is None:
                st.info(
                    "El historial no está disponible en este momento (sin conexión a la base de datos). "
                    "Podés seguir generando y descargando reportes con normalidad desde \"Cargar archivo nuevo\".",
                    icon="🗄️",
                )
            elif not filtros["aerolineas"]:
                st.info("Todavía no hay ningún reporte generado guardado en el historial.", icon="🗄️")
            else:
                col_aerolinea, col_periodo, col_estacion = st.columns(3)
                with col_aerolinea:
                    aerolinea_opciones = ["Todas"] + [a.upper() for a in filtros["aerolineas"]]
                    aerolinea_sel = st.selectbox("Aerolínea", aerolinea_opciones, key="hist_aerolinea")
                with col_periodo:
                    periodo_opciones = {"Todos": None}
                    for periodo_fecha, mes, anio in filtros["periodos"]:
                        periodo_opciones[period_label((mes, anio))] = periodo_fecha
                    periodo_sel_label = st.selectbox("Período", list(periodo_opciones.keys()), key="hist_periodo")
                with col_estacion:
                    estacion_opciones = ["Todas"] + filtros["estaciones"]
                    estacion_sel = st.selectbox("Estación", estacion_opciones, key="hist_estacion")

                aerolinea_filtro = None if aerolinea_sel == "Todas" else aerolinea_sel.lower()
                periodo_filtro = periodo_opciones[periodo_sel_label]
                estacion_filtro = None if estacion_sel == "Todas" else estacion_sel

                historial_df = obtener_historial(aerolinea_filtro, periodo_filtro, estacion_filtro)

                if historial_df is None:
                    st.warning("No se pudo consultar el historial ahora mismo. Probá de nuevo en unos minutos.", icon="⚠️")
                elif historial_df.empty:
                    st.info("No hay liquidaciones para los filtros seleccionados.", icon="🔍")
                else:
                    # El total SOLO suma la generacion mas reciente de cada
                    # combinacion (aerolinea/estacion/tipo_cargo/periodo) --
                    # es_vigente lo calcula obtener_historial() con ROW_NUMBER().
                    # Si algo se regenero, la tabla de abajo sigue mostrando
                    # todas las corridas para trazabilidad, pero el numero
                    # grande no debe duplicar plata.
                    vigentes = historial_df[historial_df["es_vigente"]]
                    total = float(vigentes["monto_total"].fillna(0).sum())
                    cantidad_vigente = len(vigentes)
                    cantidad_total = len(historial_df)

                    st.markdown(
                        f'<div class="hwc-group-card">'
                        f'<div class="hwc-group-title">Σ Total filtrado</div>'
                        f'<div class="hwc-row hwc-row-active"><span class="hwc-dot hwc-dot-active">✓</span>'
                        f'{cantidad_vigente} liquidaci{"ón" if cantidad_vigente == 1 else "ones"} vigente{"" if cantidad_vigente == 1 else "s"}'
                        f'<span class="hwc-count">{_money(total)}</span></div>'
                        f'</div>',
                        unsafe_allow_html=True,
                    )
                    if cantidad_total > cantidad_vigente:
                        st.caption(
                            f"ℹ️ Hay {cantidad_total - cantidad_vigente} generación(es) anterior(es) para "
                            "alguna combinación aerolínea/estación/tipo de cargo/período — el total de "
                            "arriba usa solo la más reciente de cada una. Las anteriores siguen abajo, "
                            "marcadas como \"Anterior\", solo para trazabilidad."
                        )

                    # ---------------------------------------------------------
                    # Dos graficos simples, nativos de Streamlit (sin libs
                    # nuevas). Ambos usan SOLO montos vigentes (misma logica
                    # que el total de arriba) para no duplicar plata.
                    # ---------------------------------------------------------
                    col_chart1, col_chart2 = st.columns(2)
                    with col_chart1:
                        st.markdown(
                            '<div class="hwc-group-title">📊 Total facturado por aerolínea</div>',
                            unsafe_allow_html=True,
                        )
                        por_aerolinea = (
                            vigentes.assign(_aerolinea=vigentes["aerolinea"].str.upper())
                            .groupby("_aerolinea")["monto_total"]
                            .sum()
                        )
                        st.altair_chart(
                            _grafico_barras(por_aerolinea, "Aerolínea", "#2D79AB"),
                            use_container_width=True,
                        )

                    with col_chart2:
                        st.markdown(
                            '<div class="hwc-group-title">📈 Evolución por mes</div>',
                            unsafe_allow_html=True,
                        )
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
                            # El orden cronologico (no alfabetico -- "julio"
                            # quedaria antes que "mayo") se fuerza pasando la
                            # lista de categorias ya ordenada como "sort" del
                            # eje X de Altair, en vez de alfabetico por
                            # defecto.
                            st.altair_chart(
                                _grafico_barras(por_mes, "Período", "#3895D1", orden=list(por_mes.index)),
                                use_container_width=True,
                            )
                            if len(por_mes) == 1:
                                st.caption("Todavía hay un solo período cargado — este gráfico va a sumar meses a medida que se generen más reportes.")

                    tabla = pd.DataFrame({
                        "Vigencia": historial_df["es_vigente"].map(lambda v: "✓ Vigente" if v else "— Anterior"),
                        "Fecha de generación": historial_df["generado_en"].dt.strftime("%d/%m/%Y %H:%M") + " UTC",
                        "Aerolínea": historial_df["aerolinea"].str.upper(),
                        "Estación": historial_df["estacion"],
                        "Tipo de cargo": historial_df["tipo_cargo"].map(
                            lambda t: CHARGE_TYPE_LABELS.get(t, ("", t))[1]
                        ),
                        "Período": [
                            period_label((mes, anio))
                            for mes, anio in zip(historial_df["periodo_mes"], historial_df["periodo_anio"])
                        ],
                        "Movimientos": historial_df["cantidad_filas"],
                        "Monto total": historial_df["monto_total"].fillna(0).map(_money),
                    })
                    st.dataframe(tabla, use_container_width=True, hide_index=True)

                    # ---------------------------------------------------------
                    # Detalle linea por linea de UNA liquidacion puntual. El
                    # combo de abajo referencia filas de historial_df por
                    # posicion (indice 0..N-1), no por un id propio -- alcanza
                    # porque se reconstruye en cada rerun a partir del mismo
                    # query, en el mismo orden.
                    # ---------------------------------------------------------
                    st.markdown(
                        '<div class="hwc-step" style="margin-top:1.4rem;">'
                        '<span class="hwc-step-num">🔎</span>Ver detalle de una liquidación</div>',
                        unsafe_allow_html=True,
                    )

                    def _etiqueta_detalle(i: int) -> str:
                        fila = historial_df.iloc[i]
                        tipo_label = CHARGE_TYPE_LABELS.get(fila["tipo_cargo"], ("", fila["tipo_cargo"]))[1]
                        periodo_txt = period_label((fila["periodo_mes"], fila["periodo_anio"]))
                        fecha_txt = fila["generado_en"].strftime("%d/%m/%Y %H:%M UTC")
                        vigencia_txt = "vigente" if fila["es_vigente"] else "anterior"
                        return (
                            f"{fila['aerolinea'].upper()} · {fila['estacion']} · {tipo_label} · "
                            f"{periodo_txt} · {fecha_txt} ({vigencia_txt})"
                        )

                    seleccion = st.selectbox(
                        "Elegí una liquidación para ver su detalle línea por línea",
                        options=range(len(historial_df)),
                        format_func=_etiqueta_detalle,
                        label_visibility="collapsed",
                        key="hist_detalle_selector",
                    )
                    fila_sel = historial_df.iloc[seleccion]

                    if fila_sel["cantidad_filas"] == 0:
                        st.info("Esta liquidación no tiene movimientos asociados (sin movimiento).", icon="🔍")
                    else:
                        with st.spinner("Cargando detalle..."):
                            resultado_detalle = _obtener_detalle_liquidacion(fila_sel)
                        if resultado_detalle is None:
                            st.warning(
                                "No se pudo cargar el detalle ahora mismo. Probá de nuevo en unos minutos.",
                                icon="⚠️",
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
                                    f"{len(detalle_df)} guías (hoja \"GUIAS\") — TC {resumen['tipo_cambio']:,.2f}. "
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
                                la_rows = "".join(
                                    f'<div class="hwc-row hwc-row-active"><span class="hwc-dot hwc-dot-active">✓</span>'
                                    f'{col}<span class="hwc-count">{_money(resumen["sums"][col])}</span></div>'
                                    for col in LATAM_SUBFACTURA_LA
                                )
                                m4_rows = "".join(
                                    f'<div class="hwc-row hwc-row-active"><span class="hwc-dot hwc-dot-active">✓</span>'
                                    f'{col}<span class="hwc-count">{_money(resumen["sums"][col])}</span></div>'
                                    for col in LATAM_SUBFACTURA_4M
                                )
                                st.markdown(
                                    f'<div class="hwc-group-card">'
                                    f'<div class="hwc-group-title">🧾 Resumen Facturación</div>'
                                    f'<div class="hwc-row" style="font-weight:700;color:var(--hwc-blue-text);">LATAM AIRLINES</div>'
                                    f'{la_rows}'
                                    f'<div class="hwc-row hwc-row-active"><span class="hwc-dot hwc-dot-active">Σ</span>'
                                    f'<b>TOTAL LA (sin IVA)</b><span class="hwc-count">{_money(resumen["total_la"])}</span></div>'
                                    f'<div class="hwc-row" style="font-weight:700;color:var(--hwc-blue-text);margin-top:0.6rem;">LAN ARGENTINA</div>'
                                    f'{m4_rows}'
                                    f'<div class="hwc-row hwc-row-active"><span class="hwc-dot hwc-dot-active">%</span>'
                                    f'IVA (21%, informativo)<span class="hwc-count">{_money(resumen["iva"])}</span></div>'
                                    f'<div class="hwc-row hwc-row-active"><span class="hwc-dot hwc-dot-active">Σ</span>'
                                    f'<b>TOTAL 4M (con IVA)</b><span class="hwc-count">{_money(resumen["total_4m"])}</span></div>'
                                    f'<div class="hwc-row hwc-row-active" style="margin-top:0.5rem;border-top:1px solid var(--hwc-border);padding-top:0.6rem;">'
                                    f'<span class="hwc-dot hwc-dot-active">✓</span><b>TOTAL PERIODO</b>'
                                    f'<span class="hwc-count">{_money(resumen["total_periodo"])}</span></div>'
                                    f'</div>',
                                    unsafe_allow_html=True,
                                )
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
                            st.caption("Descarga opcional — la liquidación ya quedó guardada.")
                            st.download_button(
                                label="⬇️ Descargar Excel",
                                data=detalle_buffer,
                                file_name=f"{fila_sel['aerolinea']}_{fila_sel['estacion']}_{periodo_slug}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                type="secondary",
                                key="hist_detalle_download",
                            )

                            # Alto dinamico hasta un tope: con liquidaciones
                            # chicas (ej. 3 filas) no deja un montón de grilla
                            # vacia; con liquidaciones grandes (ej. 1033 filas)
                            # se topea en 420px y scrollea adentro del recuadro
                            # en vez de estirar la pagina entera.
                            alto_tabla = min(420, 38 * (len(detalle_df) + 1) + 4)
                            st.dataframe(detalle_df, use_container_width=True, hide_index=True, height=alto_tabla)


st.markdown('<div class="hwc-footer">Handyway Cargo · Automatización de reportes</div>', unsafe_allow_html=True)
