"""Cantidad de vuelos internacionales que trae el Excel de Ariel.

La rendicion de Ariel trae, ademas de la hoja de ventas, una hoja de
manifiestos por estacion ("mani mane EZE", "mani mane MDZ"; el nombre cambia
de un mes a otro: "MANI MANE EZEIZA", "Mane EZE"...). De ahi sale la cantidad
de vuelos que se liquidan por "Commissions for sales - Inter":

- EZE: dos bloques (EXPO a la izquierda, IMPO a la derecha), cada uno con tres
  filas rotuladas al pie: "TOTAL VUELOS", "CON CARGA", "SIN CARGA". Se liquida
  lo que sale SIN carga: cantidad = suma de las dos filas "SIN CARGA". Se
  buscan por su rotulo, no por celda fija (la cantidad de vuelos cambia el
  largo de la hoja).
- MDZ: una fila por vuelo con su importe en USD (columna "USD" o "IMPORTE") y,
  al pie, el total. Cantidad = filas de vuelo con importe.

Reglas de este modulo:
- Solo lee: no guarda nada ni toca la base.
- NUNCA lanza excepciones por un archivo raro. Si no puede identificar la
  cantidad con certeza, no inventa nada: devuelve cantidad None y explica el
  problema para que la pantalla lo avise y la usuaria lo cargue a mano.
- Una hoja que no esta no es un problema (puede ser un Excel de otro tipo):
  esa estacion simplemente no aparece en el resultado.
- El total en USD que trae la hoja se devuelve solo como CONTROL (a veces viene
  sumado y a veces lo tipea Anita); nunca define la cantidad.
- Controles extra de EZE (solo avisan, no cambian la cantidad): se cuentan las
  filas de vuelo de cada bloque y se comparan con los totales tipeados al pie, y
  se avisa si el mismo numero de mane aparece en dos filas.
"""

import unicodedata
from dataclasses import dataclass

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from src.jetsmart_cruce import _rewind

ESTACIONES = ("EZE", "MDZ")


@dataclass(frozen=True)
class VuelosEstacion:
    estacion: str
    hoja: str
    cantidad: int | None  # None: no se pudo identificar con certeza
    origen: str | None  # "Tomado del Excel de Ariel, hoja ...: ..." (si hay cantidad)
    total_usd: float | None  # total en USD del Excel, solo para controlar
    problema: str | None  # por que no se pudo (si no hay cantidad)
    # Para mostrar el mismo desglose que la hoja de Anita (solo lectura):
    bloques: tuple[tuple[str, int], ...] = ()  # EZE: (("EXPO", 25), ("IMPO", 24)) vuelos sin carga
    celda_total: str | None = None  # donde esta el total en USD de la hoja ("B42", "G30")
    avisos_conteo: tuple[str, ...] = ()  # las filas de vuelo contadas no coinciden con los totales tipeados
    avisos_repetidos: tuple[str, ...] = ()  # mismo numero de mane en dos filas (puede ser un vuelo repetido)


def _norm(valor: object) -> str:
    texto = unicodedata.normalize("NFKD", str(valor if valor is not None else ""))
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return " ".join(texto.casefold().split())


def _numero(valor: object) -> float | None:
    if isinstance(valor, bool) or not isinstance(valor, (int, float)):
        return None
    return float(valor)


def _entero(valor: object) -> int | None:
    n = _numero(valor)
    return int(n) if n is not None and n >= 0 and n == int(n) else None


def _a_numero_mostrable(n: float) -> int | float:
    return int(n) if n == int(n) else n


def _vacia(valor: object) -> bool:
    return valor is None or str(valor).strip() == ""


def _celda(fila: int, col: int) -> str:
    """Coordenada de Excel de una celda (indices desde 0): (41, 1) -> "B42"."""
    return f"{get_column_letter(col + 1)}{fila + 1}"


def _primer_valor_a_la_derecha(fila: list, col: int) -> object:
    """Primera celda no vacia a la derecha del rotulo (None si no hay)."""
    for valor in fila[col + 1:]:
        if not _vacia(valor):
            return valor
    return None


def _hojas_de(wb, estacion: str) -> list[str]:
    return [n for n in wb.sheetnames if estacion.casefold() in _norm(n) and ("mane" in _norm(n) or "mani" in _norm(n))]


def _grilla(ws) -> list[list]:
    return [list(r) for r in ws.iter_rows(values_only=True)]


def _buscar(grilla: list[list], predicado) -> list[tuple[int, int]]:
    return [(i, j) for i, fila in enumerate(grilla) for j, v in enumerate(fila) if v is not None and predicado(_norm(v))]


def _valor_del_rotulo(grilla: list[list], i: int, j: int) -> int | None:
    return _entero(_primer_valor_a_la_derecha(grilla[i], j))


# Celdas que no son un vuelo: los titulos de la tabla y el marcador "SIN VUELOS" de los dias sin movimiento.
_NO_ES_VUELO = ("vuelo", "mane", "mani", "sin vuelos", "sin vuelo")

_NOMBRE_BLOQUE = {"EXPO": "Exportación", "IMPO": "Importación"}


def nombre_bloque(bloque: str) -> str:
    """"EXPO" -> "Exportación", "IMPO" -> "Importación" (como en la hoja de Anita)."""
    return _NOMBRE_BLOQUE.get(bloque, bloque.capitalize())


def _leer_eze(grilla: list[list]) -> tuple[list[tuple[str, int]] | None, float | None, str | None]:
    """([(bloque, vuelos sin carga)], total USD del Excel, problema)."""
    bloques, total, _celda_total, _controles, problema = _leer_eze_completo(grilla)
    return bloques, total, problema


def _leer_eze_completo(grilla: list[list]):
    """([(bloque, sin carga)], total USD, celda del total, controles por bloque, problema).

    Cada control es (bloque, columna del rotulo, fila de "TOTAL VUELOS", total tipeado, con carga tipeado)."""
    sin_carga = _buscar(grilla, lambda t: t == "sin carga")
    if len(sin_carga) != 2:
        return None, None, None, [], (
            f'Esperaba dos filas "SIN CARGA" (una por bloque, EXPO e IMPO) y encontré {len(sin_carga)}.'
        )
    bloques = []
    controles = []
    for i, j in sorted(sin_carga, key=lambda ij: ij[1]):
        sin = _valor_del_rotulo(grilla, i, j)
        # Las otras dos filas del bloque: las mas cercanas hacia arriba en la misma columna.
        total = con = fila_total = None
        nombre = ""
        for k in range(i - 1, -1, -1):
            v = grilla[k][j] if j < len(grilla[k]) else None
            t = _norm(v)
            if con is None and t == "con carga":
                con = _valor_del_rotulo(grilla, k, j)
            elif total is None and t.startswith("total vuelos"):
                total = _valor_del_rotulo(grilla, k, j)
                nombre = str(v).strip()[len("total vuelos"):].strip()
                fila_total = k
            if con is not None and total is not None:
                break
        if sin is None or con is None or total is None:
            return None, None, None, [], (
                'No pude leer los tres números del bloque ("TOTAL VUELOS", "CON CARGA", "SIN CARGA").'
            )
        if con + sin != total:
            return None, None, None, [], (
                f"El bloque {nombre or j + 1} no cierra: {con} con carga + {sin} sin carga no es {total} vuelos en total."
            )
        nombre_b = nombre.upper() or f"bloque {len(bloques) + 1}"
        bloques.append((nombre_b, sin))
        controles.append((nombre_b, j, fila_total, total, con))
    total_usd, celda = _total_usd_eze(grilla, max(i for i, _ in sin_carga))
    return bloques, total_usd, celda, controles, None


def _total_usd_eze(grilla: list[list], desde: int) -> tuple[float | None, str | None]:
    """(total en USD al pie, su celda). Puede venir rotulado ("TOTAL"), o suelto en
    una fila con una unica celda numerica. Si no hay una forma clara, no hay control."""
    rotulados = []
    for i in range(desde + 1, len(grilla)):
        for j, v in enumerate(grilla[i]):
            if v is not None and _norm(v) == "total":
                for k in range(j + 1, len(grilla[i])):
                    if not _vacia(grilla[i][k]):
                        rotulados.append((_numero(grilla[i][k]), _celda(i, k)))
                        break
    if len(rotulados) == 1:
        return rotulados[0]
    sueltos = []
    for i in range(desde + 1, len(grilla)):
        celdas = [(k, v) for k, v in enumerate(grilla[i]) if not _vacia(v)]
        if len(celdas) == 1 and _numero(celdas[0][1]) is not None:
            sueltos.append((_numero(celdas[0][1]), _celda(i, celdas[0][0])))
    return sueltos[0] if len(sueltos) == 1 else (None, None)


def _controles_eze(grilla: list[list], controles: list, hoja: str) -> tuple[tuple[str, ...], tuple[str, ...]]:
    """(avisos de conteo, avisos de mane repetido) de los bloques de EZE.

    Cuenta las filas de vuelo de cada bloque (las que tienen vuelo o numero de mane)
    y las que van con carga (guias > 0), y las compara con los totales tipeados al
    pie. Tambien busca el mismo numero de mane en dos filas."""
    conteo, repetidos = [], []
    for bloque, j, fila_total, total, con in controles:
        filas = [
            i for i in range(fila_total)
            if any(k < len(grilla[i]) and not _vacia(grilla[i][k]) and _norm(grilla[i][k]) not in _NO_ES_VUELO
                   for k in (j, j + 1))
        ]
        con_carga = [i for i in filas if j + 2 < len(grilla[i]) and (_numero(grilla[i][j + 2]) or 0) > 0]
        nombre = nombre_bloque(bloque)
        if len(filas) != total or len(con_carga) != con:
            conteo.append(
                f"{nombre}: en la hoja {hoja} conté {len(filas)} filas de vuelo ({len(con_carga)} con carga), "
                f"pero los totales de abajo dicen {total} vuelos ({con} con carga). Revisá cuál está bien."
            )
        por_mane: dict[str, list[int]] = {}
        for i in filas:
            if j + 1 < len(grilla[i]) and not _vacia(grilla[i][j + 1]):
                por_mane.setdefault(_norm(grilla[i][j + 1]), []).append(i)
        for filas_mane in por_mane.values():
            if len(filas_mane) > 1:
                mane = str(grilla[filas_mane[0]][j + 1]).strip()
                donde = ", ".join(str(i + 1) for i in filas_mane[:-1]) + f" y {filas_mane[-1] + 1}"
                repetidos.append(
                    f"{nombre}: el mane {mane} aparece en las filas {donde} de la hoja {hoja}. "
                    "Puede ser un vuelo repetido: revisalo."
                )
    return tuple(conteo), tuple(repetidos)


def _leer_mdz(grilla: list[list]) -> tuple[int | None, float | None, str | None]:
    """(vuelos con importe, total USD del Excel, problema)."""
    cantidad, total, _celda_total, problema = _leer_mdz_completo(grilla)
    return cantidad, total, problema


def _leer_mdz_completo(grilla: list[list]):
    """(vuelos con importe, total USD, celda del total, problema)."""
    for h, fila in enumerate(grilla):
        normas = [_norm(v) for v in fila]
        if "vuelo" in normas and ("usd" in normas or "importe" in normas):
            col_vuelo = normas.index("vuelo")
            col_usd = normas.index("usd") if "usd" in normas else normas.index("importe")
            break
    else:
        return None, None, None, 'No encontré las columnas "VUELO" y "USD" (o "IMPORTE").'
    vuelos = 0
    totales = []
    for n in range(h + 1, len(grilla)):
        fila = grilla[n]
        vuelo = fila[col_vuelo] if col_vuelo < len(fila) else None
        usd = _numero(fila[col_usd]) if col_usd < len(fila) else None
        if not _vacia(vuelo) and usd is not None and usd > 0:
            vuelos += 1
        elif _vacia(vuelo) and usd is not None:
            totales.append((usd, _celda(n, col_usd)))
    if vuelos == 0:
        return None, None, None, "No encontré ninguna fila de vuelo con importe en USD."
    total, celda = totales[-1] if len(totales) == 1 else (None, None)
    return vuelos, total, celda, None


def leer_vuelos_ariel(source) -> dict[str, VuelosEstacion]:
    """{estacion: VuelosEstacion} de las hojas de manifiestos que tenga el Excel.
    Un Excel sin esas hojas (o que no se puede abrir) devuelve {}."""
    try:
        _rewind(source)
        wb = load_workbook(source, data_only=True, read_only=True)
    except Exception:
        return {}
    resultado: dict[str, VuelosEstacion] = {}
    try:
        for estacion in ESTACIONES:
            hojas = _hojas_de(wb, estacion)
            if not hojas:
                continue
            nombre = wb[hojas[0]].title.strip()
            if len(hojas) > 1:
                resultado[estacion] = VuelosEstacion(
                    estacion, nombre, None, None, None,
                    f"Hay más de una hoja de manifiestos de {estacion}: " + ", ".join(h.strip() for h in hojas) + ".",
                )
                continue
            bloques, celda_total, conteo, repetidos = (), None, (), ()
            try:
                grilla = _grilla(wb[hojas[0]])
                if estacion == "EZE":
                    leidos, total_usd, celda_total, controles, problema = _leer_eze_completo(grilla)
                    cantidad = sum(n for _, n in leidos) if leidos else None
                    origen = None
                    if cantidad is not None:
                        bloques = tuple(leidos)
                        partes = " + ".join(str(n) for _, n in leidos)
                        origen = f"Tomado del Excel de Ariel, hoja {nombre}: {partes} vuelos sin carga = {cantidad}"
                        try:
                            conteo, repetidos = _controles_eze(grilla, controles, nombre)
                        except Exception:  # los controles son un extra: nunca tumban la lectura
                            conteo, repetidos = (), ()
                else:
                    cantidad, total_usd, celda_total, problema = _leer_mdz_completo(grilla)
                    origen = None
                    if cantidad is not None:
                        origen = f"Tomado del Excel de Ariel, hoja {nombre}: {cantidad} vuelos con USD"
            except Exception:
                cantidad, total_usd, origen = None, None, None
                bloques, celda_total, conteo, repetidos = (), None, (), ()
                problema = "No pude leer la hoja."
            if total_usd is not None:
                total_usd = _a_numero_mostrable(total_usd)
            resultado[estacion] = VuelosEstacion(
                estacion, nombre, cantidad, origen, total_usd, problema,
                bloques, celda_total if total_usd is not None else None, conteo, repetidos,
            )
    finally:
        wb.close()
    return resultado
