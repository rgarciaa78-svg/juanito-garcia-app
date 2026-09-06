#!/usr/bin/env python3
"""
JUANITO — Recoge las exportaciones del Analizador desde Descargas.

Power BI las baja todas con el mismo nombre (`PowerBIPerformanceData (n).json`),
así que renombrarlas a mano es tedioso y fácil de equivocar. Este script las
mueve a `capturas/` y las bautiza con los visuales que contienen, que es lo
único que identifica de verdad de qué pestaña salió cada una.

Si dos archivos traen exactamente los mismos visuales, se conserva el primero
y se avisa: exportar dos veces la misma pestaña es un error frecuente.

Uso:  python3 scripts/recoger_capturas.py
"""

import json
import os
import re
import shutil
import sys

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DESTINO = os.path.join(RAIZ, "capturas")
DESCARGAS = os.path.expanduser("~/Downloads")

# Visuales que no dicen nada de la pestaña: aparecen en casi todas.
GENERICOS = {"imagen", "segmentación de datos", "segmentacion de datos",
             "tarjeta", "forma", "botón", "boton", "(sin título)"}


def visuales(ruta):
    with open(ruta, encoding="utf-8") as fh:
        datos = json.load(fh)
    out, vistos = [], set()
    for e in datos.get("events", []):
        t = (e.get("metrics") or {}).get("visualTitle")
        if t and t not in vistos:
            vistos.add(t)
            out.append(t)
    return out


def n_consultas(ruta):
    with open(ruta, encoding="utf-8") as fh:
        datos = json.load(fh)
    return sum(1 for e in datos.get("events", [])
               if isinstance(e.get("metrics"), dict) and "QueryText" in e["metrics"])


def apodo(vs):
    """Nombre de archivo a partir de los visuales que NO son genéricos."""
    utiles = [v for v in vs if v.strip().lower() not in GENERICOS]
    base = "-".join(utiles[:2]) if utiles else "-".join(vs[:2]) or "sin-visuales"
    base = base.lower()
    base = (base.replace("á", "a").replace("é", "e").replace("í", "i")
                .replace("ó", "o").replace("ú", "u").replace("ñ", "n"))
    base = re.sub(r"[^a-z0-9]+", "-", base).strip("-")
    return base[:70] or "captura"


def main():
    os.makedirs(DESTINO, exist_ok=True)
    pendientes = sorted(
        (os.path.join(DESCARGAS, f) for f in os.listdir(DESCARGAS)
         if f.startswith("PowerBIPerformanceData") and f.endswith(".json")),
        key=os.path.getmtime)
    if not pendientes:
        print("No hay exportaciones nuevas en Descargas.")
        return 0

    # Firmas de lo que ya está guardado, para no duplicar.
    firmas = {}
    for f in os.listdir(DESTINO):
        if f.endswith(".json") and f != "catalogo.json":
            try:
                firmas["|".join(visuales(os.path.join(DESTINO, f)))] = f
            except Exception:
                pass

    movidos = 0
    for origen in pendientes:
        try:
            vs = visuales(origen)
            nq = n_consultas(origen)
        except Exception as e:
            print(f"  ✗ {os.path.basename(origen)}: no se pudo leer ({e})")
            continue
        firma = "|".join(vs)
        if firma in firmas:
            # Puede ser una exportación repetida, o que el archivo ya se haya
            # movido a mano a capturas/. En ambos casos no hay nada que hacer.
            print(f"  · {os.path.basename(origen)}: ya está en capturas/ como "
                  f"{firmas[firma]} — se omite")
            continue
        nombre = apodo(vs) + ".json"
        n = 2
        while os.path.exists(os.path.join(DESTINO, nombre)):
            nombre = f"{apodo(vs)}-{n}.json"
            n += 1
        shutil.move(origen, os.path.join(DESTINO, nombre))
        firmas[firma] = nombre
        movidos += 1
        utiles = [v for v in vs if v.strip().lower() not in GENERICOS]
        print(f"  ✓ {nombre}  ({nq} consultas) — {', '.join(utiles[:4]) or 'solo genéricos'}")

    print(f"\n{movidos} exportación(es) recogida(s). "
          f"Ahora: python3 scripts/leer_capturas.py")
    return 0


if __name__ == "__main__":
    sys.exit(main())
