"""Configuracion parametrizada del proceso de armado de reportes por aerolinea.

Todo lo que depende de la aerolinea o del tipo de cargo vive aca. Agregar una
aerolinea nueva (ej. LATAM) es sumar una entrada a AIRLINE_CONFIGS; agregar un
tipo de cargo nuevo es sumar una entrada a CHARGE_TYPES.
"""

# Columnas del "archivo original.xlsx" tal como vienen del export del sistema.
COL_CODIGO = "Código"
COL_COD_VUELO = "Cod.Vuelo"
COL_CLIENTE = "Cliente"
COL_CONDICION = "Condición"
COL_TIPO = "Tipo"
COL_TPO_CAMBIO = "Tpo.Cambio"
COL_DRY_FEE = "Dry.Fee"
COL_ADUANA = "Aduana"
COL_TRANS_E = "Trans.E."
COL_IATA = "Iata"
COL_COLLECT = "Collect"
COL_AEROLINEA = "Aerolinea"
COL_ESTACION = "Estacion"

# Estaciones "canonicas" para las que siempre se genera una hoja (con datos o
# con el texto de "sin movimiento"). NQN queda afuera: en el archivo original
# solo aparece con movimientos de LATAM, ninguna otra aerolinea opera ahi.
STATIONS = ["EZE", "COR", "ROS", "MDZ", "AEP"]

# Definicion de cada tipo de cargo: que columna de monto filtra las filas,
# que columnas se llevan al reporte final, y si se separa en una hoja por
# estacion o se agrupa en una unica hoja para todo el pais.
CHARGE_TYPES = {
    "delivery_fee": {
        "sheet_prefix": "Delivery Fee",
        "amount_columns": [COL_DRY_FEE],
        "report_columns": [COL_CODIGO, COL_COD_VUELO, COL_CONDICION, COL_TPO_CAMBIO, COL_DRY_FEE],
        "per_station": True,
    },
    "trans_electronica": {
        "sheet_prefix": "Transmis Electr",
        "amount_columns": [COL_TRANS_E],
        "report_columns": [COL_CODIGO, COL_COD_VUELO, COL_TIPO, COL_TPO_CAMBIO, COL_TRANS_E],
        "per_station": True,
    },
    "collect": {
        # El cargo Collect no se separa por estacion: se arma una unica hoja
        # "Collect AR" con la columna Estacion incluida, tal como en el
        # reporte de referencia de GOL.
        "sheet_name": "Collect AR",
        "amount_columns": [COL_IATA, COL_COLLECT],
        "report_columns": [COL_CODIGO, COL_COD_VUELO, COL_CONDICION, COL_TPO_CAMBIO, COL_IATA, COL_COLLECT, COL_ESTACION],
        "per_station": False,
    },
}

# Como tratar una hoja sin ninguna fila (estacion sin movimiento, o tipo de
# cargo agrupado sin datos): "placeholder_text" arma una hoja con una unica
# celda de texto explicando que no hay movimiento; "headers_only" arma la
# hoja vacia, solo con los encabezados de columna, sin ningun texto.
EMPTY_SHEET_STYLE_PLACEHOLDER = "placeholder_text"
EMPTY_SHEET_STYLE_HEADERS_ONLY = "headers_only"

# Que flujo de armado de reporte usa cada aerolinea. AV y Gol usan el
# reporte simple (una hoja por tipo de cargo/estacion, ver report_builder.py
# y CHARGE_TYPES). LATAM usa una liquidacion formal con 3 hojas (Resumen
# Facturacion, Detalle de facturacion, Compensacion) que no encaja en ese
# mismo esquema -- ver liquidacion_builder.py.
FLOW_SIMPLE_REPORT = "simple_report"
FLOW_LIQUIDACION = "liquidacion"

# Que tipos de cargo se facturan a cada aerolinea. Esto es una decision de
# negocio, no algo que se pueda inferir de los datos: por ejemplo GOL tiene
# filas con Trans.E. != 0 en el archivo original pero ese cargo no se le
# reporta (se lo gestiona por otra via).
AIRLINE_CONFIGS = {
    "avianca": {
        "match": "AVIANCA",
        "flow": FLOW_SIMPLE_REPORT,
        "charge_types": ["delivery_fee", "trans_electronica"],
        # Sin confirmar todavia si Avianca quiere el mismo criterio que Gol
        # (hoja vacia sin texto). Hasta que lo confirmen, mantiene el texto
        # explicativo. Para cambiarlo: EMPTY_SHEET_STYLE_HEADERS_ONLY.
        "empty_sheet_style": EMPTY_SHEET_STYLE_PLACEHOLDER,
    },
    "gol": {
        "match": "GOL",
        "flow": FLOW_SIMPLE_REPORT,
        "charge_types": ["delivery_fee", "collect"],
        # Confirmado por Cristian (cliente): para Gol, las hojas sin
        # movimiento quedan vacias (solo encabezados), sin el texto de
        # "Sin movimiento de awbs...".
        "empty_sheet_style": EMPTY_SHEET_STYLE_HEADERS_ONLY,
    },
    "latam": {
        "match": "LATAM",
        # LATAM no usa el reporte simple: usa la liquidacion formal de
        # liquidacion_builder.py (ver LATAM_* mas abajo). No hay charge_types
        # ni empty_sheet_style aca porque esas opciones son del otro flujo.
        "flow": FLOW_LIQUIDACION,
    },
}

EMPTY_STATION_TEXT = "Sin movimiento de awbs de importación destino {station}"

# Delivery Fee de Avianca: regla de negocio fija, sin tabla de tarifas,
# confirmada estacion por estacion (no es una unica regla para toda
# Avianca). Confirmado por Cristian Nagel que EZE usa 200 USD fijo (margen
# de Handyway). COR, ROS, MDZ, AEP: pendiente de confirmar si aplica la
# misma regla u otra distinta -- por ahora se usa el dato bruto del
# original, que coincide con la referencia conocida para COR.
# None = sin ajuste, se deja "Dry.Fee" tal cual viene del original.
AVIANCA_DELIVERY_FEE_USD_POR_ESTACION = {
    "EZE": 200,
    "COR": None,
    "ROS": None,
    "MDZ": None,
    "AEP": None,
}

# --- Liquidacion formal de LATAM (ver liquidacion_builder.py) ---
# El archivo de referencia real que confirmo el cliente
# ("7.Liquidación Cliente GHA EZE Julio 2026.xlsx") es especificamente para
# la estacion EZE. No hay evidencia ni referencia para otras estaciones de
# LATAM (COR, ROS, MDZ, NQN) con este formato de liquidacion, asi que por
# ahora el flujo se limita a EZE.
LATAM_LIQUIDACION_STATION = "EZE"

# Nombres largos de columna para la hoja "Detalle de facturacion", tal como
# los espera el cliente. Salen directo de una columna del archivo original
# cada uno, sin ningun ajuste de tarifa -- LATAM usa el dato bruto tal cual,
# igual que Gol (a diferencia de Avianca en EZE).
COL_LATAM_DELIVERY_FEE = "Cargo Administrativo Guías Aéreas (Delivery Fee)"
COL_LATAM_ADUANA = "Recupero Gastos de Aduana RG3244/11"
COL_LATAM_EXPEDICION = "Expedición Regimen pago en Destino (2% Flete Collect o Min USD 10)"
COL_LATAM_COLLECT = "Cargos Collect AWB - Flete Internacional de Importación"

# Mapea cada columna de monto del original a su nombre largo en el Detalle
# de facturacion, en el orden en que se muestran en esa hoja.
LATAM_DETALLE_CHARGE_COLUMNS = {
    COL_DRY_FEE: COL_LATAM_DELIVERY_FEE,
    COL_ADUANA: COL_LATAM_ADUANA,
    COL_IATA: COL_LATAM_EXPEDICION,
    COL_COLLECT: COL_LATAM_COLLECT,
}

# Como se agrupan esas 4 columnas en las dos sub-facturas del Resumen
# Facturacion. "LATAM AIRLINES" (Collect AWB + Recupero Aduana) va sin IVA;
# "LAN ARGENTINA" (Delivery Fee + Expedicion) lleva 21% de IVA agregado.
# Formula verificada al centavo contra los numeros reales del archivo de
# referencia del cliente.
#
# OJO: en el archivo de referencia, el total de Collect AWB que entra a
# "LATAM AIRLINES" esta neto de un ajuste manual de "Compensacion" de ese
# periodo (ver la hoja Compensacion). Como esa compensacion es un caso a
# caso que el cliente carga a mano (no sale de "archivo original.xlsx"),
# esta liquidacion generada automaticamente usa el monto de Collect AWB en
# bruto -- va a diferir del historico por el monto de compensacion que
# haya aplicado ese mes, hasta que se cargue la hoja Compensacion a mano.
LATAM_SUBFACTURA_LA = [COL_LATAM_COLLECT, COL_LATAM_ADUANA]
LATAM_SUBFACTURA_4M = [COL_LATAM_DELIVERY_FEE, COL_LATAM_EXPEDICION]
LATAM_IVA_RATE = 0.21
