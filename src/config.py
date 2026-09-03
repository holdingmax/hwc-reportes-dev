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

# Que tipos de cargo se facturan a cada aerolinea. Esto es una decision de
# negocio, no algo que se pueda inferir de los datos: por ejemplo GOL tiene
# filas con Trans.E. != 0 en el archivo original pero ese cargo no se le
# reporta (se lo gestiona por otra via).
AIRLINE_CONFIGS = {
    "avianca": {
        "match": "AVIANCA",
        "charge_types": ["delivery_fee", "trans_electronica"],
    },
    "gol": {
        "match": "GOL",
        "charge_types": ["delivery_fee", "collect"],
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
    # },
}

EMPTY_STATION_TEXT = "Sin movimiento de awbs de importación destino {station}"
