"""Guias de JetSmart "para revisar", en palabras de Anita.

El cruce con Ariel deja una lista de excepciones (jetsmart_cruce.py) con un
tipo y un detalle tecnico. Para la pantalla de resultado cada tipo se explica
en lenguaje simple: que le paso a la guia y que tiene que hacer. Solo
presentacion: no calcula ni cambia nada del cruce ni de la liquidacion.
"""

import pandas as pd

from src.jetsmart_cruce import (
    EXC_DIFERENCIA,
    EXC_DOBLE,
    EXC_DUPLICADA,
    EXC_PENDIENTE,
    EXC_REVISION,
    nro_guia_entero,
)

# tipo -> (nombre corto para el resumen, que le paso, que hacer). El orden es el
# de aparicion: primero lo que mas probablemente haya que mirar.
_TIPOS = {
    EXC_PENDIENTE: (
        ("pendiente de declarar", "pendientes de declarar"),
        "Está en el sistema, pero Ariel todavía no la declaró.",
        "No entra este mes. Esperá a que Ariel la declare (probablemente el mes que viene).",
    ),
    EXC_REVISION: (
        ("para revisar a mano", "para revisar a mano"),
        "Ariel la declaró, pero no está en el sistema y no se pudo confirmar.",
        "No entra. Fijate si la guía existe en el sistema antes de cerrar el mes.",
    ),
    EXC_DUPLICADA: (
        ("con el número repetido", "con el número repetido"),
        "El mismo número de guía aparece más de una vez en la rendición de Ariel.",
        "Pedile a Ariel que revise cuál es el número correcto.",
    ),
    EXC_DIFERENCIA: (
        ("con diferencias", "con diferencias"),
        "El sistema y Ariel no coinciden",
        "Ya entra con los valores de Ariel. Confirmá con Ariel que sean los correctos.",
    ),
    EXC_DOBLE: (
        ("ya declarada el mes anterior", "ya declaradas el mes anterior"),
        "Ariel ya la declaró el mes anterior.",
        "No entra, para no liquidarla dos veces. No hay que hacer nada, salvo que sea un error.",
    ),
}


def _entra(texto: object) -> str:
    """"Sí, con valores de Ariel" -> "Sí (valores de Ariel)"; el resto, tal cual ("No")."""
    texto = str(texto or "").strip()
    if texto.startswith("Sí"):
        return "Sí (valores de Ariel)" if "Ariel" in texto else "Sí"
    return texto or "No"


def filas_para_revisar(excepciones: pd.DataFrame) -> list[dict]:
    """Una fila por guia a revisar: {guia (entero), que_pasa, entra, que_hacer}.
    Un tipo que esta funcion no conoce se muestra con el detalle tecnico del
    cruce, para no esconder nunca una guia."""
    if excepciones.empty:
        return []
    orden = {tipo: i for i, tipo in enumerate(_TIPOS)}
    ordenadas = excepciones.assign(
        _orden=excepciones["Tipo"].map(lambda t: orden.get(t, len(orden))),
        _guia=nro_guia_entero(excepciones["Nro guía"]),
    ).sort_values(["_orden", "_guia"], kind="stable")
    filas = []
    for _, fila in ordenadas.iterrows():
        tipo = fila["Tipo"]
        detalle = "" if pd.isna(fila["Detalle"]) else str(fila["Detalle"])
        if tipo in _TIPOS:
            _, que_pasa, que_hacer = _TIPOS[tipo]
            if tipo == EXC_DIFERENCIA:
                que = detalle.replace("Difiere ", "").rstrip(".")
                que_pasa = f"{que_pasa}: {que}." if que else f"{que_pasa}."
        else:
            que_pasa, que_hacer = f"{tipo}. {detalle}".strip(), "Revisala a mano."
        filas.append({
            "guia": None if pd.isna(fila["_guia"]) else int(fila["_guia"]),
            "que_pasa": que_pasa,
            "entra": _entra(fila["¿Entra a la liquidación?"]),
            "que_hacer": que_hacer,
        })
    return filas


def resumen_por_tipo(excepciones: pd.DataFrame) -> list[str]:
    """["7 pendientes de declarar", "1 con diferencias"], en el orden de _TIPOS."""
    if excepciones.empty:
        return []
    cuentas = excepciones["Tipo"].value_counts()
    textos = []
    for tipo, ((singular, plural), _, _) in _TIPOS.items():
        n = int(cuentas.get(tipo, 0))
        if n:
            textos.append(f"{n} {singular if n == 1 else plural}")
    desconocidos = int(sum(n for t, n in cuentas.items() if t not in _TIPOS))
    if desconocidos:
        textos.append(f"{desconocidos} otras")
    return textos
