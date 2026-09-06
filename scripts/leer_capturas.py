#!/usr/bin/env python3
"""
JUANITO — Catálogo de consultas DAX exportadas del Analizador de rendimiento.

El Analizador de Power BI tiene un botón "Exportar" que vuelca TODAS las
consultas de una pestaña en un solo archivo. Eso reemplaza el ir y venir de
"Copiar consulta" visual por visual: una exportación por pestaña en vez de una
consulta por mensaje.

Este script lee los archivos que haya en `capturas/`, saca de cada uno el DAX
de cada visual con su nombre, y arma un catálogo en
`capturas/catalogo.json`. También avisa cuáles ya están implementadas,
comparando contra las consultas que los scripts construyen.

Cómo se identifica el visual: el evento que trae el `QueryText` cuelga, por
`parentId`, de un "Visual Container Lifecycle" que sí tiene `visualTitle`. Se
sube por esa cadena hasta encontrarlo.

Uso:  python3 scripts/leer_capturas.py
"""

import hashlib
import json
import os
import re
import sys

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DIR = os.path.join(RAIZ, "capturas")
SALIDA = os.path.join(DIR, "catalogo.json")


def titulo_del_visual(evento, por_id):
    """Sube por la cadena de padres hasta el contenedor que tiene título."""
    actual = evento
    vistos = set()
    while actual is not None and actual.get("id") not in vistos:
        vistos.add(actual.get("id"))
        titulo = (actual.get("metrics") or {}).get("visualTitle")
        if titulo:
            return titulo, (actual.get("metrics") or {}).get("visualType", "")
        actual = por_id.get(actual.get("parentId"))
    return "(sin título)", ""


def consultas_de(ruta):
    with open(ruta, encoding="utf-8") as fh:
        datos = json.load(fh)
    eventos = datos.get("events", datos if isinstance(datos, list) else [])
    por_id = {e["id"]: e for e in eventos if e.get("id")}
    salida = []
    for e in eventos:
        m = e.get("metrics") or {}
        dax = m.get("QueryText")
        if not dax:
            continue
        titulo, tipo = titulo_del_visual(e, por_id)
        salida.append({
            "visual": titulo,
            "tipo_visual": tipo,
            "filas": m.get("RowCount"),
            "dax": dax,
            # El hash permite detectar duplicados entre pestañas: el mismo
            # visual repetido no debería implementarse dos veces.
            "hash": hashlib.sha1(dax.encode("utf-8")).hexdigest()[:12],
        })
    return salida


def medidas_y_tablas(dax):
    """Extrae los nombres de tabla y medida que aparecen en la consulta.

    Sirve para decidir qué aporta cada visual sin leer el DAX entero.
    """
    tablas = set(re.findall(r"'([^']+)'\[", dax))
    campos = set(re.findall(r"'[^']+'\[([^\]]+)\]", dax))
    return sorted(tablas), sorted(campos)


def ya_implementada(dax, fuentes):
    """¿Alguna parte distintiva de esta consulta ya está en los scripts?

    Se compara por medidas: si todas las medidas del visual ya aparecen en el
    código, es muy probable que ya esté cubierto. Es una pista, no un
    veredicto — por eso se reporta como 'probable'.
    """
    _t, campos = medidas_y_tablas(dax)
    if not campos:
        return False
    presentes = sum(1 for c in campos if c in fuentes)
    return presentes == len(campos)


def main():
    if not os.path.isdir(DIR):
        print(f"No existe {DIR}")
        return 1
    archivos = sorted(f for f in os.listdir(DIR)
                      if f.lower().endswith(".json") and f != "catalogo.json")
    if not archivos:
        print(f"No hay exportaciones en {DIR}. Ver capturas/LEEME.md")
        return 1

    fuentes = ""
    for s in ("fetch_powerbi.py", "fetch_series.py"):
        with open(os.path.join(RAIZ, "scripts", s), encoding="utf-8") as fh:
            fuentes += fh.read()

    catalogo, vistos = [], {}
    for nombre in archivos:
        for c in consultas_de(os.path.join(DIR, nombre)):
            c["archivo"] = nombre
            if c["hash"] in vistos:
                c["duplicado_de"] = vistos[c["hash"]]
            else:
                vistos[c["hash"]] = f"{nombre} / {c['visual']}"
            c["tablas"], c["campos"] = medidas_y_tablas(c["dax"])
            c["probablemente_implementada"] = ya_implementada(c["dax"], fuentes)
            catalogo.append(c)

    with open(SALIDA, "w", encoding="utf-8") as fh:
        json.dump(catalogo, fh, ensure_ascii=False, indent=2)

    nuevas = [c for c in catalogo
              if not c.get("duplicado_de") and not c["probablemente_implementada"]]
    print(f"{len(archivos)} exportación(es) · {len(catalogo)} consultas")
    print(f"  ya cubiertas (probable): "
          f"{sum(1 for c in catalogo if c['probablemente_implementada'])}")
    print(f"  duplicadas entre pestañas: "
          f"{sum(1 for c in catalogo if c.get('duplicado_de'))}")
    print(f"  PENDIENTES de implementar: {len(nuevas)}")
    print()
    for c in nuevas:
        print(f"  [{c['archivo']}] {c['visual']}  ({c['tipo_visual']}, "
              f"{c['filas']} filas)")
        print(f"      campos: {', '.join(c['campos'][:8])}"
              f"{' …' if len(c['campos']) > 8 else ''}")
    print()
    print(f"Catálogo completo en {SALIDA}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
