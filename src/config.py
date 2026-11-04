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
# JetSmart usa su propia liquidacion (ver jetsmart_builder.py): parte de otro
# export (el del sistema de guias, no "archivo original.xlsx") y su
# estructura (Ventas / Comisiones / GHA Services / Total a WCS) no tiene
# nada en comun con las sub-facturas de LATAM.
FLOW_JETSMART = "jetsmart_liquidacion"

# Que tipos de cargo se facturan a cada aerolinea. Esto es una decision de
# negocio, no algo que se pueda inferir de los datos: por ejemplo GOL tiene
# filas con Trans.E. != 0 en el archivo original pero ese cargo no se le
# reporta (se lo gestiona por otra via).
# Estaciones donde el dato que se muestra HOY no esta validado contra
# ningun archivo de referencia real. No se descartan ni se bloquean (ver
# nota junto a cada caso en report_builder.py / liquidacion_builder.py):
# es una red de seguridad para que la interfaz avise en vez de mostrar un
# numero sin confirmar con total confianza. Facil de vaciar (lista vacia)
# a medida que el cliente confirme cada estacion.
LATAM_UNCONFIRMED_STATIONS = ["COR", "ROS", "MDZ", "NQN"]

# Estaciones donde Avianca factura cada tipo de cargo. Confirmado por
# Cristian (cliente) el 17/9/2026: esto NO es una tarifa pendiente de
# confirmar, es una regla de negocio real -- estas combinaciones de
# estacion/tipo de cargo directamente no existen. Por eso no se genera
# hoja (ni vacia con aviso de "sin movimiento") para las estaciones que
# quedan afuera de cada lista; ver "charge_type_stations" mas abajo y
# STATIONS en report_builder.py. Delivery Fee y Transmision Electronica
# no son simetricos a proposito: Delivery Fee se factura en EZE y COR,
# pero Transmision Electronica solo en EZE, aunque el archivo original
# traiga filas de Trans.E. != 0 en otras estaciones (esas filas no se le
# facturan a Avianca por esta via).
AVIANCA_DELIVERY_FEE_STATIONS = ["EZE", "COR"]
AVIANCA_TRANS_ELECTRONICA_STATIONS = ["EZE"]

AIRLINE_CONFIGS = {
    "avianca": {
        "match": "AVIANCA",
        "flow": FLOW_SIMPLE_REPORT,
        "charge_types": ["delivery_fee", "trans_electronica"],
        # Sin confirmar todavia si Avianca quiere el mismo criterio que Gol
        # (hoja vacia sin texto). Hasta que lo confirmen, mantiene el texto
        # explicativo. Para cambiarlo: EMPTY_SHEET_STYLE_HEADERS_ONLY.
        "empty_sheet_style": EMPTY_SHEET_STYLE_PLACEHOLDER,
        # Restringe, por tipo de cargo, la lista global STATIONS a las
        # estaciones donde ese cargo realmente se factura (ver comentario de
        # AVIANCA_DELIVERY_FEE_STATIONS / AVIANCA_TRANS_ELECTRONICA_STATIONS
        # arriba). Un tipo de cargo sin entrada aca usaria STATIONS entera.
        "charge_type_stations": {
            "delivery_fee": AVIANCA_DELIVERY_FEE_STATIONS,
            "trans_electronica": AVIANCA_TRANS_ELECTRONICA_STATIONS,
        },
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
        "unconfirmed_stations": LATAM_UNCONFIRMED_STATIONS,
    },
    "jetsmart": {
        # Valor de la columna "Empresa" del export de guias (ver
        # COL_JS_EMPRESA mas abajo); se compara sin distinguir mayusculas.
        "match": "JETSMART",
        "flow": FLOW_JETSMART,
        # Un unico tipo de cargo: el flete domestico ($ Prioridad), que es
        # la base de Ventas Netas de la liquidacion. Ver JETSMART_* abajo.
        "charge_type": "jetsmart_liquidacion",
    },
}

EMPTY_STATION_TEXT = "Sin movimiento de awbs de importación destino {station}"

# Delivery Fee de Avianca: regla de negocio fija, sin tabla de tarifas,
# confirmada estacion por estacion (no es una unica regla para toda
# Avianca). Confirmado por Cristian Nagel que EZE usa 200 USD fijo (margen
# de Handyway). COR: pendiente de confirmar si aplica la misma regla u
# otra distinta -- por ahora se usa el dato bruto del original, que
# coincide con la referencia conocida para COR. ROS, MDZ y AEP no
# aparecen aca porque Avianca no factura Delivery Fee en esas estaciones
# (ver AVIANCA_DELIVERY_FEE_STATIONS mas arriba): no hay hoja donde
# aplicar ningun ajuste.
# None = sin ajuste, se deja "Dry.Fee" tal cual viene del original.
AVIANCA_DELIVERY_FEE_USD_POR_ESTACION = {
    "EZE": 200,
    "COR": None,
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

# --- Liquidacion de JetSmart (ver jetsmart_builder.py) ---
# Columnas del export mensual del sistema de guias de Handyway (NO es
# "archivo original.xlsx": es otro reporte, con una fila por guia). Nombres
# tal cual vienen en el export real de julio 2026 -- ojo con "Aero. Dstino"
# y "$ Total." (asi, con el typo y el punto).
COL_JS_CREACION = "Creacion"
COL_JS_GUIA = "# Guía"
COL_JS_CLIENTE = "Cliente"
COL_JS_ORIGEN = "Aero. Origen"
COL_JS_DESTINO = "Aero. Dstino"
COL_JS_ESTADO = "Estado"
COL_JS_KGS = "KGs"
COL_JS_PRIORIDAD = "$ Prioridad"
COL_JS_EMPRESA = "Empresa"

# Estaciones de la tabla "Ex / Kg / Rate / TC / Pesos" de la hoja
# LIQUIDACION, en el mismo orden que la liquidacion armada a mano
# (LIQ_ECS_07-2026). Los Kg se agrupan por ESTACION DE ORIGEN de la guia.
# Una estacion que aparezca en el export y no este aca se agrega al final
# de la tabla (ver build_jetsmart_resumen): nunca se descartan kilos.
JETSMART_STATIONS = ["AEP", "BRC", "COR", "CRD", "IGR", "MDZ", "NQN", "TUC", "SLA", "USH", "EZE"]

# Tarifa GHA en USD por kg, por estacion. Confirmado por Anita (1/10/2026):
# 0,185 USD/kg fija en todas las estaciones, solo cambia si cambia el
# contrato. Una estacion sin entrada aca usa JETSMART_TARIFA_USD_KG_DEFAULT.
JETSMART_TARIFA_USD_KG_DEFAULT = 0.185
JETSMART_TARIFA_USD_KG_POR_ESTACION = {station: 0.185 for station in JETSMART_STATIONS}
# Estaciones cuya tarifa esta confirmada: todas las de JETSMART_STATIONS (la
# tarifa fija es del contrato de JetSmart en general, no por estacion). Una
# estacion nueva que aparezca en el export y no este en esa lista se calcula
# igual con el default, pero muestra el aviso amarillo de "tarifa no
# confirmada" hasta que se la sume aca.
JETSMART_TARIFA_CONFIRMADA_STATIONS = ["AEP", "BRC", "COR", "CRD", "IGR", "MDZ", "NQN", "TUC", "SLA", "USH", "EZE"]

# Porcentajes de la hoja LIQUIDACION, relevados de LIQ_ECS_07-2026 (formulas
# de la planilla de Anita: B7 = -B6*0.075; IVA por guia = gravado*0.21).
JETSMART_IVA_RATE = 0.21
JETSMART_COMISION_DOMESTICA_RATE = 0.075
# "IVA of Service And Comissions" e "IIBB Tax (Over cost)" estan en 0 fijo
# en la liquidacion de julio 2026 (no es formula, es un 0 tipeado). Se
# mantienen asi hasta que se confirme si alguna vez llevan valor.
JETSMART_IVA_SERVICIOS = 0.0
JETSMART_IIBB = 0.0

# "Commissions for sales - Inter": USD fijo por vuelo internacional, por
# estacion (sale de las hojas "MANI EZE" / "MANI MDZ" de la liquidacion a
# mano: EZE 17,5 USD por vuelo, MDZ 35 USD por vuelo). La CANTIDAD de
# vuelos sale de los manifiestos, que no vienen en ningun export: se carga a
# mano al generar la liquidacion, igual que el tipo de cambio.
JETSMART_COMISION_INTER_USD_POR_VUELO = {"EZE": 17.5, "MDZ": 35.0}
# Por encima de esta cantidad de vuelos por estacion la pantalla avisa (no
# bloquea): suele ser el importe tipeado en el campo de la cantidad.
JETSMART_VUELOS_INTER_AVISO_CANTIDAD = 150
