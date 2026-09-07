#!/usr/bin/env python3
"""
JUANITO — Verificación estática, sin tocar Power BI.

Nació porque tres corridas seguidas del workflow devolvieron datos vacíos por
un error que era detectable sin salir de la máquina: un bloque de código que
leía `found["__top_proveedores"]` había quedado dentro de build_compras en vez
de build_cxp, así que nadie leía lo que la consulta sí estaba trayendo. No
había error que registrar, solo un dato que se perdía en el camino.

Correr esto ANTES de pedir una corrida del workflow. Verifica:

  1. Que cada clave interna "__algo" que alguien escribe, alguien la lea.
  2. Que cada dax_* y serie_* definida se llame desde algún sitio.
  3. Que las funciones de construcción produzcan los KPIs esperados.
  4. Que los nombres de campo que emite el script sean los que lee juanito.html.

Uso:  python3 scripts/verificar.py
"""

import importlib.util
import os
import re
import sys

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FALLOS = []


def revisar(ok, mensaje):
    print(f"  {'OK ' if ok else 'FALLA'}  {mensaje}")
    if not ok:
        FALLOS.append(mensaje)


def cargar_fetch_powerbi():
    """Importa el script con credenciales falsas: solo se usan en get_token()."""
    for var in ("AZURE_TENANT_ID", "AZURE_CLIENT_ID", "AZURE_CLIENT_SECRET",
                "PBI_USERNAME", "PBI_PASSWORD"):
        os.environ.setdefault(var, "verificacion")
    ruta = os.path.join(RAIZ, "scripts", "fetch_powerbi.py")
    spec = importlib.util.spec_from_file_location("fp", ruta)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def leer(*partes):
    with open(os.path.join(RAIZ, *partes), encoding="utf-8") as fh:
        return fh.read()


def funcion_contenedora(lineas, idx):
    for j in range(idx, -1, -1):
        if lineas[j].startswith("def "):
            return lineas[j][4:lineas[j].index("(")]
    return "?"


def test_claves_internas(src):
    """Toda clave '__x' escrita debe leerse, y viceversa."""
    print("\n1. Claves internas escritas vs leídas")
    lineas = src.split("\n")
    escrituras, lecturas = {}, {}
    for i, linea in enumerate(lineas):
        for m in re.finditer(r'\["(__[a-z_]+)"\]\s*=', linea):
            escrituras.setdefault(m.group(1), []).append(funcion_contenedora(lineas, i))
        for m in re.finditer(r'found\.get\("(__[a-z_]+)"\)', linea):
            lecturas.setdefault(m.group(1), []).append(funcion_contenedora(lineas, i))
    for clave in sorted(set(escrituras) | set(lecturas)):
        e, l = escrituras.get(clave, []), lecturas.get(clave, [])
        revisar(bool(e) and bool(l),
                f"{clave}: escribe {e or 'NADIE'} — lee {l or 'NADIE'}")


def test_orden_carga_build(src):
    """Cada bloque que trae datos debe correr ANTES de construir su reporte.

    El margen por cliente salió vacío tres corridas seguidas por esto: la
    consulta funcionaba y guardaba el resultado en `scanned`, pero
    build_margen() ya se había ejecutado veinte líneas antes. No hay error
    ni diagnóstico posible — el dato llega tarde y nadie lo lee.
    """
    print("\n5. Orden: la carga de datos antes de construir el reporte")
    lineas = src.split("\n")

    def linea_de(texto):
        for i, l in enumerate(lineas):
            if texto in l:
                return i + 1
        return None

    pares = [
        ("Margen por cliente (pestaña", "r, ventas, margen = build_margen"),
        ("Compras: faltantes y necesidad", "r, ratio = build_compras"),
        ("CxP: top 15 proveedores", "r, dias = build_cxp"),
        ("CxC: aging real", "r, sem, razon, mora_pct, vencer = build_cxc"),
        ("S&OP: clasificación del inventario", "r = build_inventario"),
        ("Avance vs Presupuesto: tabla por canal", "av, avance_pct = build_avance"),
    ]
    for carga, build in pares:
        a, b = linea_de(carga), linea_de(build)
        if a is None or b is None:
            revisar(False, f"no se encontró '{carga[:34]}' o su build")
            continue
        revisar(a < b, f"{carga[:38]}: carga L{a} → build L{b}")


def test_funciones_usadas(src, prefijo, archivo):
    print(f"\n2. Funciones {prefijo}* de {archivo} que nadie llama")
    for nombre in re.findall(rf"^def ({prefijo}[a-z0-9_]+)\(", src, re.M):
        llamadas = len(re.findall(rf"\b{nombre}\(", src)) - 1
        revisar(llamadas > 0, f"{nombre} ({llamadas} llamada/s)")


def test_construccion(fp):
    """Cada build_* con datos sintéticos debe producir sus KPIs y estructuras."""
    print("\n3. Funciones de construcción con datos sintéticos")

    casos = [
        ("build_cxc",
         lambda: fp.build_cxc({
             "% Morosidad": 0.2108,
             "__aging": [("1. Por Vencer", 3570000.0), ("2. 0 a 15 días", 226121.0),
                         ("3. 16 a 30 días", 277775.0), ("4. Más de 30 días", 442248.0)]}),
         ["Morosidad", "CxC por vencer", "CxC Total", "CxC Vencido"], ["tramos"]),
        ("build_cxp",
         lambda: fp.build_cxp({
             "Cuentas x Pagar": 16810000, "Refinanciamiento": 6960000, "Proveedores": 139,
             "__top_proveedores": [("E & M S.R.L.", "LECHE", 5175254.0)],
             "__aging_cxp": [("Vigente", 6467432.0), ("0 a 7 días", 600000.0),
                             ("8 a 15 días", 454715.0), ("31 a 90 días", 1300005.0)]}),
         ["CxP Total", "Refinanciado", "# Proveedores", "Deuda top 15 proveedores",
          "CxP Vencido"],
         ["proveedores_criticos", "tramos"]),
        ("build_fill_rate",
         lambda: fp.build_fill_rate({
             "__fillrate_card": 0.884, "__venta_mes": 1720.0, "__no_atendido_total": 199.0,
             "__por_grupo": [("SPSA", 0.925)], "__soles_por_grupo": {"SPSA": 129.0},
             "__por_marca": [("MAQUILA", 118.0)]}),
         ["Fill Rate", "Venta del mes (tarjeta)", "Pedidos no atendidos (mes)"],
         ["por_grupo", "por_marca"]),
        ("build_inventario",
         lambda: fp.build_inventario({"__clasificacion": [
             ("ENVASES Y EMBALAJES", "1. Working", 1900000.0, None),
             ("ENVASES Y EMBALAJES", "4. Dead", 917801.0, None),
             ("PT SALSAS", "2. Exceso 1", 714044.0, None)]}),
         ["Inventario Total", "Dead Stock", "% Dead Stock", "Working Stock"],
         ["dead_por_categoria", "working_por_categoria"]),
    ]
    salidas = {}
    for nombre, fn, kpis_esperados, estructuras in casos:
        res = fn()
        if isinstance(res, tuple):
            res = res[0]
        salidas.update({k: v for k, v in res.items() if k in estructuras})
        etiquetas = [k["label"] for k in res.get("kpis", [])]
        faltan = [k for k in kpis_esperados if k not in etiquetas]
        revisar(not faltan, f"{nombre}: KPIs {etiquetas}"
                            + (f" — FALTAN {faltan}" if faltan else ""))
        for est in estructuras:
            revisar(bool(res.get(est)), f"{nombre}: produce '{est}'")
    return salidas


def test_campos(salidas, html):
    """Los campos emitidos deben ser los que juanito.html lee."""
    print("\n4. Campos emitidos vs campos que lee la app")
    variable = {"tramos": "t", "proveedores_criticos": "p", "dead_por_categoria": "c",
                "working_por_categoria": "c", "por_grupo": "g", "por_marca": "m"}
    ignorar = {"map", "join", "length", "filter", "slice", "forEach"}
    for estructura, var in variable.items():
        filas = salidas.get(estructura)
        if not filas:
            revisar(False, f"{estructura}: no se produjo, no se puede comparar")
            continue
        emitidos = set(filas[0].keys())
        pos = html.find(f"?.{estructura} || []")
        if pos < 0:
            pos = html.find(f".{estructura} || []")
        bloque = html[pos:pos + 1500] if pos >= 0 else ""
        leidos = {u for u in re.findall(rf"\b{var}\.([a-zA-Z_]+)\b", bloque)} - ignorar
        huerfanos = sorted(leidos - emitidos)
        revisar(not huerfanos,
                f"{estructura}: la app lee {sorted(leidos)}"
                + (f" — NO EXISTEN {huerfanos}" if huerfanos else ""))


def main():
    src_pbi = leer("scripts", "fetch_powerbi.py")
    src_ser = leer("scripts", "fetch_series.py")
    html = leer("juanito.html")

    test_claves_internas(src_pbi)
    test_orden_carga_build(src_pbi)
    test_funciones_usadas(src_pbi, "dax_", "fetch_powerbi.py")
    test_funciones_usadas(src_ser, "serie_", "fetch_series.py")
    salidas = test_construccion(cargar_fetch_powerbi())
    test_campos(salidas, html)

    print()
    if FALLOS:
        print(f"{len(FALLOS)} problema(s) — corregir ANTES de correr el workflow:")
        for f in FALLOS:
            print(f"  · {f}")
        return 1
    print("Todo conectado. El workflow puede correrse.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
