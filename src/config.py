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

# Que tipos de cargo se facturan a cada aerolinea. Esto es una decision de
# negocio, no algo que se pueda inferir de los datos: por ejemplo GOL tiene
# filas con Trans.E. != 0 en el archivo original pero ese cargo no se le
# reporta (se lo gestiona por otra via).
AIRLINE_CONFIGS = {
    "avianca": {
        "match": "AVIANCA",
        "charge_types": ["delivery_fee", "trans_electronica"],
        # Sin confirmar todavia si Avianca quiere el mismo criterio que Gol
        # (hoja vacia sin texto). Hasta que lo confirmen, mantiene el texto
        # explicativo. Para cambiarlo: EMPTY_SHEET_STYLE_HEADERS_ONLY.
        "empty_sheet_style": EMPTY_SHEET_STYLE_PLACEHOLDER,
    },
    "gol": {
        "match": "GOL",
        "charge_types": ["delivery_fee", "collect"],
        # Confirmado por Cristian (cliente): para Gol, las hojas sin
        # movimiento quedan vacias (solo encabezados), sin el texto de
        # "Sin movimiento de awbs...".
        "empty_sheet_style": EMPTY_SHEET_STYLE_HEADERS_ONLY,
    },
    # LATAM: PENDIENTE DE CONFIRMAR CON EL CLIENTE.
    # No hay un reporte de referencia armado a mano para LATAM (como si hay
    # para AV y GOL), asi que no se puede inferir con la misma confianza que
    # tipos de cargo se le facturan realmente. El ejemplo de abajo asume los
    # tres tipos de cargo porque en "archivo original.xlsx" LATAM es la unica
    # aerolinea (junto con GOL) que tiene filas != 0 en Dry.Fee, Trans.E. e
    # Iata/Collect al mismo tiempo -- pero eso es solo evidencia de que hay
    # datos, no confirmacion de que Handyway le facture ese cargo (recordar
    # el caso de GOL, que tiene Trans.E. != 0 en los datos pero NO se le
    # reporta ese cargo). Para activarla: descomentar y ajustar charge_types
    # segun lo que confirme el cliente.
    # "latam": {
    #     "match": "LATAM",
    #     "charge_types": ["delivery_fee", "trans_electronica", "collect"],
    #     "empty_sheet_style": EMPTY_SHEET_STYLE_PLACEHOLDER,
    # },
}

EMPTY_STATION_TEXT = "Sin movimiento de awbs de importación destino {station}"
