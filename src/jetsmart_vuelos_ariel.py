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
"""

import unicodedata
from dataclasses import dataclass

from openpyxl import load_workbook

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


def _primer_valor_a_la_derecha(fila: list, col: int) -> object:
    """Primera celda no vacia a la derecha del rotulo (None si no hay)."""
    for valor in fila[col + 1:]:
        if valor is not None and str(valor).strip() != "":
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


def _leer_eze(grilla: list[list]) -> tuple[list[tuple[str, int]] | None, float | None, str | None]:
    """([(bloque, vuelos sin carga)], total USD del Excel, problema)."""
    sin_carga = _buscar(grilla, lambda t: t == "sin carga")
    if len(sin_carga) != 2:
        return None, None, (
            f'Esperaba dos filas "SIN CARGA" (una por bloque, EXPO e IMPO) y encontré {len(sin_carga)}.'
        )
    bloques = []
    for i, j in sorted(sin_carga, key=lambda ij: ij[1]):
        sin = _valor_del_rotulo(grilla, i, j)
        # Las otras dos filas del bloque: las mas cercanas hacia arriba en la misma columna.
        total = con = None
        nombre = ""
        for k in range(i - 1, -1, -1):
            v = grilla[k][j] if j < len(grilla[k]) else None
            t = _norm(v)
            if con is None and t == "con carga":
                con = _valor_del_rotulo(grilla, k, j)
            elif total is None and t.startswith("total vuelos"):
                total = _valor_del_rotulo(grilla, k, j)
                nombre = str(v).strip()[len("total vuelos"):].strip()
            if con is not None and total is not None:
                break
        if sin is None or con is None or total is None:
            return None, None, (
                'No pude leer los tres números del bloque ("TOTAL VUELOS", "CON CARGA", "SIN CARGA").'
            )
        if con + sin != total:
            return None, None, (
                f"El bloque {nombre or j + 1} no cierra: {con} con carga + {sin} sin carga no es {total} vuelos en total."
            )
        bloques.append((nombre.upper() or f"bloque {len(bloques) + 1}", sin))
    return bloques, _total_usd_eze(grilla, max(i for i, _ in sin_carga)), None


def _total_usd_eze(grilla: list[list], desde: int) -> float | None:
    """Total en USD al pie. Puede venir rotulado ("TOTAL"), o suelto en una fila
    con una unica celda numerica. Si no hay una forma clara, no hay control."""
    abajo = grilla[desde + 1:]
    rotulados = [
        _numero(_primer_valor_a_la_derecha(f, j))
        for f in abajo for j, v in enumerate(f) if v is not None and _norm(v) == "total"
    ]
    if len(rotulados) == 1:
        return rotulados[0]
    sueltos = []
    for fila in abajo:
        celdas = [v for v in fila if v is not None and str(v).strip() != ""]
        if len(celdas) == 1 and _numero(celdas[0]) is not None:
            sueltos.append(_numero(celdas[0]))
    return sueltos[0] if len(sueltos) == 1 else None


def _leer_mdz(grilla: list[list]) -> tuple[int | None, float | None, str | None]:
    """(vuelos con importe, total USD del Excel, problema)."""
    for h, fila in enumerate(grilla):
        normas = [_norm(v) for v in fila]
        if "vuelo" in normas and ("usd" in normas or "importe" in normas):
            col_vuelo = normas.index("vuelo")
            col_usd = normas.index("usd") if "usd" in normas else normas.index("importe")
            break
    else:
        return None, None, 'No encontré las columnas "VUELO" y "USD" (o "IMPORTE").'
    vuelos = 0
    totales = []
    for fila in grilla[h + 1:]:
        vuelo = fila[col_vuelo] if col_vuelo < len(fila) else None
        usd = _numero(fila[col_usd]) if col_usd < len(fila) else None
        hay_vuelo = vuelo is not None and str(vuelo).strip() != ""
        if hay_vuelo and usd is not None and usd > 0:
            vuelos += 1
        elif not hay_vuelo and usd is not None:
            totales.append(usd)
    if vuelos == 0:
        return None, None, "No encontré ninguna fila de vuelo con importe en USD."
    return vuelos, (totales[-1] if len(totales) == 1 else None), None


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
            try:
                grilla = _grilla(wb[hojas[0]])
                if estacion == "EZE":
                    bloques, total_usd, problema = _leer_eze(grilla)
                    cantidad = sum(n for _, n in bloques) if bloques else None
                    origen = None
                    if cantidad is not None:
                        partes = " + ".join(str(n) for _, n in bloques)
                        origen = f"Tomado del Excel de Ariel, hoja {nombre}: {partes} vuelos sin carga = {cantidad}"
                else:
                    cantidad, total_usd, problema = _leer_mdz(grilla)
                    origen = None
                    if cantidad is not None:
                        origen = f"Tomado del Excel de Ariel, hoja {nombre}: {cantidad} vuelos con USD"
            except Exception:
                cantidad, total_usd, origen = None, None, None
                problema = "No pude leer la hoja."
            if total_usd is not None:
                total_usd = _a_numero_mostrable(total_usd)
            resultado[estacion] = VuelosEstacion(estacion, nombre, cantidad, origen, total_usd, problema)
    finally:
        wb.close()
    return resultado
