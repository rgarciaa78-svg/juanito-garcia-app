#!/usr/bin/env python3
"""
JUANITO — Serie histórica mensual desde Power BI
Extrae el valor de cada medida confirmada mes a mes desde Ene-2025 hasta el mes actual.
Salida: data/latest/series.json

Estructura:
{
  "generado": "2026-09-02",
  "periodos": ["2025-01", "2025-02", ...],
  "datasets": {
     "cxc": { "% Morosidad": [0.14, 0.15, ...], ... },
     ...
  }
}

Todos los valores provienen exclusivamente de la API REST de Power BI (executeQueries).
"""

import os, json, re, time, datetime, requests
from pathlib import Path

TENANT_ID     = os.environ["AZURE_TENANT_ID"]
CLIENT_ID     = os.environ["AZURE_CLIENT_ID"]
CLIENT_SECRET = os.environ["AZURE_CLIENT_SECRET"]
USERNAME      = os.environ["PBI_USERNAME"]
PASSWORD      = os.environ["PBI_PASSWORD"]

PBI_SCOPE = "https://analysis.windows.net/powerbi/api/.default"
TOKEN_URL = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
PBI_BASE  = "https://api.powerbi.com/v1.0/myorg"
WS_ID     = "461932ad-b5ec-4fd6-aa97-f1fc7bdc5169"

OUTPUT_DIR = Path("data/latest")
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

DATASET_IDS = {
    "cxc":            "2eec70cd-0820-408f-b938-a2cd547b0c18",
    "cxp":            "45a8ab8d-e162-4398-a260-3a9a5f90829f",
    "margen":         "38076daa-d2cd-4a93-858a-82c0a4cf8cb6",
    "mermas":         "35866214-f4da-45a3-a5a2-aa0c8caffe78",
    "compras":        "06408938-8202-424e-80c0-b42c178dabde",
    "inventario":     "0e27d784-41a4-48f0-9208-60210119f0a7",
    "control_ds":     "0aca7bdd-6b72-49c2-be41-17ae0f6b5848",
    "planificacion":  "30074d92-7ec1-4762-82f2-1cb29c15dcfe",
    "consumo":        "c972c8cb-e5fc-4b60-8f5e-265a78e1e796",
    "fill_rate":      "7f4ebe22-5e90-4e35-973b-4af3c58497e5",
}

# Pares (tabla, columna) de fecha a probar, en orden de probabilidad
DATE_CANDIDATES = [
    ("Calendario", "Date"),
    ("Calendario", "Fecha"),
    ("Calendario", "FECHA"),
    ("Fecha", "Date"),
    ("Fecha", "Fecha"),
    ("Calendar", "Date"),
    ("DimCalendario", "Date"),
    ("Dim Calendario", "Date"),
    ("dCalendario", "Date"),
    ("Tiempo", "Fecha"),
]

INICIO_ANIO = 2025
INICIO_MES  = 1


def get_token():
    r = requests.post(TOKEN_URL, data={
        "grant_type": "password", "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET, "username": USERNAME,
        "password": PASSWORD, "scope": PBI_SCOPE,
    }, timeout=30)
    r.raise_for_status()
    return r.json()["access_token"]


def dax(token, dataset_id, query, label="q", retries=3):
    url = f"{PBI_BASE}/groups/{WS_ID}/datasets/{dataset_id}/executeQueries"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    body = {"queries": [{"query": query}], "serializerSettings": {"includeNulls": True}}
    for attempt in range(retries):
        try:
            r = requests.post(url, json=body, headers=headers, timeout=60)
            if r.status_code == 429:
                wait = int(r.headers.get("Retry-After", 12))
                print(f"      [{label}] throttled — espera {wait}s")
                time.sleep(wait)
                continue
            if r.status_code != 200:
                if attempt == retries - 1:
                    print(f"      [{label}] HTTP {r.status_code}: {r.text[:150]}")
                return None
            tables = r.json().get("results", [{}])[0].get("tables", [])
            return tables[0].get("rows", []) if tables else []
        except Exception as e:
            if attempt == retries - 1:
                print(f"      [{label}] error: {e}")
            else:
                time.sleep(3)
    return None


def _sanear_no_finitos(obj):
    """Reemplaza Infinity/-Infinity/NaN por None dentro de listas y dicts.

    Devuelve cuántos reemplazó. Necesario porque json.dumps los escribiría
    literalmente y el JSON dejaría de ser parseable por el navegador.
    """
    import math
    n = 0
    if isinstance(obj, dict):
        it = obj.items()
    elif isinstance(obj, list):
        it = enumerate(obj)
    else:
        return 0
    for k, v in it:
        if isinstance(v, float) and not math.isfinite(v):
            obj[k] = None
            n += 1
        else:
            n += _sanear_no_finitos(v)
    return n


def build_periodos():
    """Lista de (anio, mes) desde Ene-2025 hasta el mes anterior al actual."""
    hoy = datetime.date.today()
    fin_anio, fin_mes = (hoy.year, hoy.month - 1) if hoy.month > 1 else (hoy.year - 1, 12)
    out, a, m = [], INICIO_ANIO, INICIO_MES
    while (a, m) <= (fin_anio, fin_mes):
        out.append((a, m))
        m += 1
        if m > 12:
            m, a = 1, a + 1
    return out


def _probar_par(token, dataset_id, medida, tbl, col):
    q = (f'EVALUATE ROW("v", CALCULATE([{medida}], '
         f"FILTER(ALL('{tbl}'), YEAR('{tbl}'[{col}]) = 2025 && MONTH('{tbl}'[{col}]) = 6)))")
    rows = dax(token, dataset_id, q, f"probe {tbl}[{col}]", retries=1)
    return rows is not None and len(rows) > 0


def columnas_fecha_reales(token, dataset_id):
    """Descubre columnas de tipo fecha del modelo vía DMV INFO.COLUMNS()/INFO.TABLES()."""
    tablas = {}
    rows = dax(token, dataset_id, "EVALUATE INFO.TABLES()", "info-tables", retries=1)
    if rows:
        for r in rows:
            tid = r.get("[ID]", r.get("ID"))
            nm  = r.get("[Name]", r.get("Name"))
            if tid is not None and nm:
                tablas[tid] = nm

    cols = dax(token, dataset_id, "EVALUATE INFO.COLUMNS()", "info-columns", retries=1)
    pares = []
    if cols:
        for c in cols:
            nm = c.get("[ExplicitName]", c.get("ExplicitName")) or c.get("[Name]", c.get("Name"))
            dt = str(c.get("[ExplicitDataType]", c.get("ExplicitDataType", "")) or "")
            tid = c.get("[TableID]", c.get("TableID"))
            tname = tablas.get(tid)
            if not nm or not tname:
                continue
            # 9 = DateTime en el modelo tabular; además acepta nombres típicos
            es_fecha = dt == "9" or any(k in nm.lower() for k in ("fecha", "date", "dia", "día"))
            if es_fecha:
                pares.append((tname, nm))
    return pares


def detectar_fecha(token, dataset_id, medida):
    """Encuentra la tabla/columna de fecha que sí filtra la medida."""
    # 1) Nombres frecuentes (rápido)
    for tbl, col in DATE_CANDIDATES:
        if _probar_par(token, dataset_id, medida, tbl, col):
            print(f"    Tabla de fecha: '{tbl}'[{col}]")
            return tbl, col

    # 2) Descubrimiento real del modelo
    print("    Candidatos comunes fallaron — leyendo esquema del modelo...")
    pares = columnas_fecha_reales(token, dataset_id)
    if pares:
        print(f"    Columnas de fecha halladas: {pares[:12]}")
    vistos = set(DATE_CANDIDATES)
    for tbl, col in pares:
        if (tbl, col) in vistos:
            continue
        vistos.add((tbl, col))
        if _probar_par(token, dataset_id, medida, tbl, col):
            print(f"    Tabla de fecha: '{tbl}'[{col}]")
            return tbl, col
    return None, None


def serie_batch(token, dataset_id, medidas, tbl, col, periodos):
    """Un solo query DAX que devuelve todos los meses y todas las medidas."""
    a_ini, m_ini = periodos[0]
    a_fin, m_fin = periodos[-1]
    ult_dia = 31
    cols = []
    for i, med in enumerate(medidas):
        cols.append(
            f'    "m{i}", VAR a = [Anio] VAR m = [Mes] RETURN '
            f"CALCULATE([{med}], FILTER(ALL('{tbl}'), "
            f"YEAR('{tbl}'[{col}]) = a && MONTH('{tbl}'[{col}]) = m))"
        )
    cols_txt = ",\n".join(cols)
    q = f"""EVALUATE
VAR Periodos =
    DISTINCT(
        SELECTCOLUMNS(
            FILTER(
                ALL('{tbl}'),
                '{tbl}'[{col}] >= DATE({a_ini}, {m_ini}, 1)
                && '{tbl}'[{col}] <= DATE({a_fin}, {m_fin}, {ult_dia})
            ),
            "Anio", YEAR('{tbl}'[{col}]),
            "Mes",  MONTH('{tbl}'[{col}])
        )
    )
RETURN
ADDCOLUMNS(
    Periodos,
{cols_txt}
)"""
    rows = dax(token, dataset_id, q, "serie-batch")
    if not rows:
        return None

    # Indexa resultados por (anio, mes)
    por_periodo = {}
    for row in rows:
        a = row.get("[Anio]", row.get("Anio"))
        m = row.get("[Mes]",  row.get("Mes"))
        if a is None or m is None:
            continue
        por_periodo[(int(a), int(m))] = row

    salida = {med: [] for med in medidas}
    for (a, m) in periodos:
        row = por_periodo.get((a, m))
        for i, med in enumerate(medidas):
            v = row.get(f"[m{i}]", row.get(f"m{i}")) if row else None
            salida[med].append(v)
    return salida


def serie_mes_a_mes(token, dataset_id, medidas, tbl, col, periodos):
    """Fallback: un query por mes con todas las medidas en un ROW()."""
    salida = {med: [] for med in medidas}
    for (a, m) in periodos:
        partes = []
        for i, med in enumerate(medidas):
            partes.append(
                f'  "m{i}", CALCULATE([{med}], FILTER(ALL(\'{tbl}\'), '
                f"YEAR('{tbl}'[{col}]) = {a} && MONTH('{tbl}'[{col}]) = {m}))"
            )
        q = "EVALUATE\nROW(\n" + ",\n".join(partes) + "\n)"
        rows = dax(token, dataset_id, q, f"{a}-{m:02d}")
        row = rows[0] if rows else None
        for i, med in enumerate(medidas):
            v = row.get(f"[m{i}]", row.get(f"m{i}")) if row else None
            salida[med].append(v)
        print(f"      {a}-{m:02d} ok")
    return salida


CXC_DATE_TBL = "LocalDateTable_9de71e0e-3dad-40db-b49a-7124064ba537"

MESES_CORTOS = {
    "ene": 1, "feb": 2, "mar": 3, "abr": 4, "may": 5, "jun": 6,
    "jul": 7, "ago": 8, "sep": 9, "set": 9, "oct": 10, "nov": 11, "dic": 12,
}


def serie_cxc_morosidad(token, ds_id, periodos):
    """Serie mensual de % Morosidad — consulta exacta del gráfico "% DE MOROSIDAD".

    Capturada con Copiar consulta el 2026-09-06. Dos hallazgos que hacían
    imposible obtener esta serie con la sonda genérica:

    1. La medida del gráfico es '[% MOROSIDAD x mes]', NO '[% MOROSIDAD]'.
       La segunda es un escalar que ignora la fecha — por eso la tarjeta
       devolvía siempre el mismo número y la serie salía "plana".
    2. El eje de tiempo no es una tabla Calendario normal: el año viene de
       una tabla de fechas automática ('LocalDateTable_9de71e0e-...') y el
       mes de 'Calendario (FACT)'. Ninguna de las dos estaba en
       DATE_CANDIDATES, y la detección automática no podía adivinarlas.

    Se envía la consulta tal cual la genera Power BI (sin reescribirla) y
    después se mapean las filas a los períodos del dashboard.
    """
    q = (
        "DEFINE\n"
        "\tVAR __DS0FilterTable = \n"
        "\t\tFILTER(\n"
        f"\t\t\tKEEPFILTERS(VALUES('{CXC_DATE_TBL}'[Año])),\n"
        f"\t\t\t'{CXC_DATE_TBL}'[Año] > 2024\n"
        "\t\t)\n\n"
        "\tVAR __DS0Core = \n"
        "\t\tSUMMARIZECOLUMNS(\n"
        f"\t\t\t'{CXC_DATE_TBL}'[Año],\n"
        "\t\t\t'Calendario (FACT)'[Mes Corto],\n"
        "\t\t\t'Calendario (FACT)'[Mes Numero],\n"
        "\t\t\t__DS0FilterTable,\n"
        "\t\t\t\"v__MOROSIDAD_x_mes\", 'DATA_FACTURACION'[% MOROSIDAD x mes]\n"
        "\t\t)\n\n"
        "EVALUATE\n\t__DS0Core\n\n"
        "ORDER BY\n"
        f"\t'{CXC_DATE_TBL}'[Año],\n"
        "\t'Calendario (FACT)'[Mes Numero]"
    )
    rows = dax(token, ds_id, q, "cxc-morosidad-mensual")
    if not rows:
        return None

    por_periodo = {}
    for r in rows:
        anio = r.get(f"{CXC_DATE_TBL}[Año]") or r.get("[Año]")
        mes = r.get("Calendario (FACT)[Mes Numero]") or r.get("[Mes Numero]")
        val = r.get("[v__MOROSIDAD_x_mes]") or r.get("v__MOROSIDAD_x_mes")
        if mes is None:
            corto = str(r.get("Calendario (FACT)[Mes Corto]") or "").strip().lower()[:3]
            mes = MESES_CORTOS.get(corto)
        if anio is None or mes is None or val is None:
            continue
        try:
            por_periodo[(int(anio), int(mes))] = float(val)
        except (TypeError, ValueError):
            continue

    if not por_periodo:
        return None
    return [por_periodo.get(p) for p in periodos]


MERMAS_MEDIDAS = [
    ("v__Merma_total",             "% Merma Total"),
    ("SumDesviacion_Bases",        "Desviación Bases"),
    ("CANT_REAL_CONSUMO_INTERNO",  "Consumo Interno"),
    ("CANT_REAL_DESTRUCCION",      "Destrucción"),
]


# Plantas. Cada entrada: (etiqueta, valor de [almacen], medida).
#
# Agrupar por [almacen] con la medida base NO sirve: el mapeo de almacén era
# correcto, pero cada planta usa su propia medida. ATE usa '% Merma total' y
# Pachacamac usa '% Merma total SALSAS' — con la base, Pachacamac daba cifras
# que no coincidían con el reporte. Tercera vez que generalizar produce
# números plausibles y equivocados, así que solo entra lo capturado.
#
# Falta PLANTA TERCEROS: sin su Copiar consulta no se sabe ni su almacén ni
# su medida, y no se van a deducir.
# Plantas. Cada entrada: (etiqueta, [almacenes], medida, lleva_TIPO_DE_BASE).
#
# Nada de esto es deducible — los cuatro visuales de planta de un mismo
# reporte resultaron distintos entre sí:
#
#   · ATE y Pachacamac: 1 almacén, 7 filtros (con TIPO DE BASE), y medidas
#     DISTINTAS entre ellas ('% Merma total' vs '% Merma total SALSAS').
#   · Terceros: 3 almacenes agrupados, 6 filtros (SIN TIPO DE BASE).
#
# Y el nombre del visual no dice nada del almacén: "PLANTA ATE" es
# "Lacteos Producción" y "PLANTA TERCEROS" son tres maquilas.
#
# Nota sobre '% Merma total MAQUILA': se llegó a descartar por no coincidir
# con el gráfico de la unidad de negocio MAQUILA (que usa MMAQUILA, con
# doble M). Sí se usa — es la de Terceros. Dos medidas de nombre casi igual
# para cosas distintas.
PLANTAS_MERMAS = [
    # etiqueta,           almacenes,                    medida,                  tipo_base
    ("Planta ATE",        ["Lacteos Producción"],       "% Merma total",         True),
    ("Planta Pachacamac", ["Salsas Producción"],        "% Merma total SALSAS",  True),
    ("Planta Terceros",   ["Abuela Maquila",
                           "Piamonte Maquila",
                           "Lacteos Dosimetria"],       "% Merma total MAQUILA", False),
]


def _q_mermas_planta(anio, almacenes, medida, alias, tipo_base=True):
    """Merma mensual de UNA planta: filtro de [almacen] + su medida propia.

    Reproduce las consultas capturadas de PLANTA ATE, PLANTA PACHACAMAC y
    PLANTA TERCEROS (Copiar consulta, 2026-09-06). Parte del bloque de
    filtros del segmento (5 ó 6 según `tipo_base`) y le añade el TREATAS de
    almacén como último VAR.
    """
    base = _q_mermas_segmento(anio, "Tabla Mermas", medida, alias, tipo_base)
    n_ultimo = 6 if tipo_base else 5          # cuántos VAR trae ya el bloque
    nuevo = n_ultimo + 1
    # Las llaves del conjunto son obligatorias: TREATAS({"a","b"}, col).
    # Sin ellas DAX interpreta cada valor como un argumento suelto y falla.
    vals = ",\n".join(f'\t\t\t\t"{a}"' for a in almacenes).lstrip("\t")
    filtro = (f"\tVAR __DS0FilterTable{nuevo} = \n"
              f"\t\tTREATAS(\n\t\t\t{{{vals}}},\n"
              f"\t\t\t'Tabla Mermas'[almacen]\n\t\t)\n\n")
    base = base.replace("\tVAR __DS0Core = \n", filtro + "\tVAR __DS0Core = \n")
    return base.replace(
        f"\t\t\t__DS0FilterTable{n_ultimo},\n",
        f"\t\t\t__DS0FilterTable{n_ultimo},\n\t\t\t__DS0FilterTable{nuevo},\n",
    )


def _q_mermas(anio):
    """Consulta del gráfico "PAUNO" de MERMA MENSUAL POR PLANTA, tal cual la
    genera Power BI (Copiar consulta, 2026-09-06), con el año del segmentador
    como parámetro.

    Los 6 filtros del visual se reproducen literalmente. Son los que hacen
    que el número sea comparable con lo que ve el equipo en el reporte:
      1. 'Calendario'[Date] >= 31/07/2025 (ventana móvil del visual)
      2. 'Tabla Mermas'[TIPO DE BASE] no vacío
      3. 'Calendario'[Año] = <anio>  ← único parámetro; equivale a mover el
         segmentador ANO del reporte, que es justamente lo que haría un
         usuario para ver 2025 en vez de 2026.
      4. categoria_producto excluye BONIFICACION Y REBATES, CHATARRA,
         INTERESES, MATERIA PRIMA, SERVICIOS, SUMINISTROS y vacío
      5. producto excluye 2 ítems puntuales
      6. estado distinto de "Cancelado"

    Sin estos filtros el % de merma sale inflado: incluye chatarra, materia
    prima, servicios y comprobantes anulados.
    """
    return (
        "DEFINE\n"
        "\tVAR __DS0FilterTable = \n"
        "\t\tFILTER(\n"
        "\t\t\tKEEPFILTERS(VALUES('Calendario'[Date])),\n"
        "\t\t\t'Calendario'[Date] >= (DATE(2025, 7, 31) + TIME(0, 0, 1))\n"
        "\t\t)\n\n"
        "\tVAR __DS0FilterTable2 = \n"
        "\t\tFILTER(\n"
        "\t\t\tKEEPFILTERS(VALUES('Tabla Mermas'[TIPO DE BASE])),\n"
        "\t\t\tNOT('Tabla Mermas'[TIPO DE BASE] IN {BLANK()})\n"
        "\t\t)\n\n"
        "\tVAR __DS0FilterTable3 = \n"
        f"\t\tTREATAS({{{anio}}}, 'Calendario'[Año])\n\n"
        "\tVAR __DS0FilterTable4 = \n"
        "\t\tFILTER(\n"
        "\t\t\tKEEPFILTERS(VALUES('Maestra de Facturacion (Total)'[categoria_producto])),\n"
        "\t\t\tNOT(\n"
        "\t\t\t\t'Maestra de Facturacion (Total)'[categoria_producto] IN {\"BONIFICACION Y REBATES\",\n"
        "\t\t\t\t\t\"CHATARRA\",\n\t\t\t\t\t\"INTERESES\",\n\t\t\t\t\t\"MATERIA PRIMA\",\n"
        "\t\t\t\t\t\"SERVICIOS\",\n\t\t\t\t\t\"SUMINISTROS\",\n\t\t\t\t\tBLANK()}\n"
        "\t\t\t)\n\t\t)\n\n"
        "\tVAR __DS0FilterTable5 = \n"
        "\t\tFILTER(\n"
        "\t\t\tKEEPFILTERS(VALUES('Maestra de Facturacion (Total)'[producto])),\n"
        "\t\t\tNOT(\n"
        "\t\t\t\t'Maestra de Facturacion (Total)'[producto] IN {\"PAVO C/M C/ASA EP CONG (8 KG)\",\n"
        "\t\t\t\t\t\"ALIMENTACION COMERCIAL\"}\n"
        "\t\t\t)\n\t\t)\n\n"
        "\tVAR __DS0FilterTable6 = \n"
        "\t\tFILTER(\n"
        "\t\t\tKEEPFILTERS(VALUES('Maestra de Facturacion (Total)'[estado])),\n"
        "\t\t\tNOT('Maestra de Facturacion (Total)'[estado] IN {\"Cancelado\"})\n"
        "\t\t)\n\n"
        "\tVAR __DS0Core = \n"
        "\t\tSUMMARIZECOLUMNS(\n"
        "\t\t\t'Calendario'[Año],\n\t\t\t'Calendario'[MES],\n"
        "\t\t\t__DS0FilterTable,\n\t\t\t__DS0FilterTable2,\n\t\t\t__DS0FilterTable3,\n"
        "\t\t\t__DS0FilterTable4,\n\t\t\t__DS0FilterTable5,\n\t\t\t__DS0FilterTable6,\n"
        "\t\t\t\"SumDesviacion_Bases\", IGNORE(CALCULATE(SUM('Tabla Mermas'[Desviacion Bases]))),\n"
        "\t\t\t\"CANT_REAL_CONSUMO_INTERNO\", IGNORE('Maestra de Kardex (Total)'[CANT REAL CONSUMO INTERNO]),\n"
        "\t\t\t\"CANT_REAL_DESTRUCCION\", IGNORE('Maestra de Kardex (Total)'[CANT REAL DESTRUCCION]),\n"
        "\t\t\t\"v__Merma_total\", 'Tabla Mermas'[% Merma total]\n"
        "\t\t)\n\n"
        "EVALUATE\n\t__DS0Core\n\n"
        "ORDER BY\n\t'Calendario'[Año], 'Calendario'[MES]"
    )


# Segmentos de merma. Cada entrada: (etiqueta, medida, lleva_filtro_tipo_base)
#
# El sondeo por nombre se retiró: encontró '% Merma total MAQUILA', que EXISTE
# en el modelo pero NO es la del gráfico — el visual usa
# '% Merma total MMAQUILA' (doble M). Sus valores no coincidían con el reporte.
# Con dos medidas de nombre casi idéntico, adivinar por el título del visual
# no es fiable, así que solo entra lo confirmado.
#
# Tampoco hay un patrón uniforme de filtros: B&D usa 5 (sin TIPO DE BASE) y
# MAQUILA usa 6 (con TIPO DE BASE). Por eso el flag es por segmento.
SEGMENTOS_MERMAS = [
    # etiqueta,   medida,                    tipo_base   origen
    ("B&D",       "% Merma total B&D",       False),   # Copiar consulta 2026-09-06
    ("MAQUILA",   "% Merma total MMAQUILA",  True),    # Copiar consulta 2026-09-06
    ("TIGO",      "% Merma total TIGO",      False),   # sondeo + valores validados
                                                       # contra el gráfico (meses 2-8)
]


def _q_mermas_segmento(anio, medida_tbl, medida, alias, tipo_base=False):
    """Consulta del gráfico "B&D" de MERMA MENSUAL POR UNIDAD DE NEGOCIO,
    tal cual la genera Power BI (Copiar consulta, 2026-09-06).

    El bloque de filtros varía según el visual: B&D usa 5 (sin TIPO DE BASE)
    y MAQUILA usa 6 (con él). `tipo_base` selecciona cuál, y la numeración de
    los VAR se ajusta sola. No es un patrón deducible — cada uno viene de su
    Copiar consulta.

    La medida también cambia por segmento y se pasa explícita: los nombres
    están en SEGMENTOS_MERMAS, ninguno se adivina.
    """
    # Numeración de los VAR: sin TIPO DE BASE van 5 filtros (patrón B&D);
    # con él van 6 y todo corre un número (patrón MAQUILA). Se respeta el
    # orden que genera Power BI en cada caso.
    n = 1
    tb = ""
    if tipo_base:
        tb = ("\tVAR __DS0FilterTable2 = \n"
              "\t\tFILTER(\n"
              "\t\t\tKEEPFILTERS(VALUES('Tabla Mermas'[TIPO DE BASE])),\n"
              "\t\t\tNOT('Tabla Mermas'[TIPO DE BASE] IN {BLANK()})\n"
              "\t\t)\n\n")
        n = 2
    v = lambda i: f"__DS0FilterTable{'' if i == 1 else i}"
    usados = "".join(f"\t\t\t{v(i)},\n" for i in range(1, n + 5))
    return (
        "DEFINE\n"
        "\tVAR __DS0FilterTable = \n"
        "\t\tFILTER(\n"
        "\t\t\tKEEPFILTERS(VALUES('Calendario'[Date])),\n"
        "\t\t\t'Calendario'[Date] >= (DATE(2025, 7, 31) + TIME(0, 0, 1))\n"
        "\t\t)\n\n"
        + tb +
        f"\tVAR {v(n + 1)} = \n"
        f"\t\tTREATAS({{{anio}}}, 'Calendario'[Año])\n\n"
        f"\tVAR {v(n + 2)} = \n"
        "\t\tFILTER(\n"
        "\t\t\tKEEPFILTERS(VALUES('Maestra de Facturacion (Total)'[categoria_producto])),\n"
        "\t\t\tNOT(\n"
        "\t\t\t\t'Maestra de Facturacion (Total)'[categoria_producto] IN {\"BONIFICACION Y REBATES\",\n"
        "\t\t\t\t\t\"CHATARRA\",\n\t\t\t\t\t\"INTERESES\",\n\t\t\t\t\t\"MATERIA PRIMA\",\n"
        "\t\t\t\t\t\"SERVICIOS\",\n\t\t\t\t\t\"SUMINISTROS\",\n\t\t\t\t\tBLANK()}\n"
        "\t\t\t)\n\t\t)\n\n"
        f"\tVAR {v(n + 3)} = \n"
        "\t\tFILTER(\n"
        "\t\t\tKEEPFILTERS(VALUES('Maestra de Facturacion (Total)'[producto])),\n"
        "\t\t\tNOT(\n"
        "\t\t\t\t'Maestra de Facturacion (Total)'[producto] IN {\"PAVO C/M C/ASA EP CONG (8 KG)\",\n"
        "\t\t\t\t\t\"ALIMENTACION COMERCIAL\"}\n"
        "\t\t\t)\n\t\t)\n\n"
        f"\tVAR {v(n + 4)} = \n"
        "\t\tFILTER(\n"
        "\t\t\tKEEPFILTERS(VALUES('Maestra de Facturacion (Total)'[estado])),\n"
        "\t\t\tNOT('Maestra de Facturacion (Total)'[estado] IN {\"Cancelado\"})\n"
        "\t\t)\n\n"
        "\tVAR __DS0Core = \n"
        "\t\tSUMMARIZECOLUMNS(\n"
        "\t\t\t'Calendario'[Año],\n\t\t\t'Calendario'[MES],\n"
        + usados +
        f"\t\t\t\"{alias}\", '{medida_tbl}'[{medida}]\n"
        "\t\t)\n\n"
        "EVALUATE\n\t__DS0Core\n\n"
        "ORDER BY\n\t'Calendario'[Año], 'Calendario'[MES]"
    )


MARGEN_LOCALDATE = "LocalDateTable_17f9bf1b-53ab-4dc5-bca5-92ec1d35bc75"

# Bloque de filtros del visual "MARGEN VARIABLE PAUNO", tal cual lo genera
# Power BI (Copiar consulta, 2026-09-06). Son ONCE filtros sobre CUATRO
# tablas distintas: 'Exl A Maestra de Facturas de Venta', 'Maestra de
# Facturacion (Total)', 'Maestra de Kardex (Total)' y 'OrdenxFactura'.
#
# Ojo: las exclusiones NO son las mismas en cada tabla. Por ejemplo
# 'Exl A...' excluye {CHATARRA, SERVICIOS} y 'Maestra de Facturacion'
# excluye {BONIFICACION Y REBATES, INTERESES, MATERIA PRIMA, SUMINISTROS,
# SERVICIOS} — CHATARRA no está en esa segunda lista. Copiar una sobre otra
# cambiaría el resultado, así que va literal.
#
# __ANIO__ es el único parámetro (el segmentador ANO del reporte).
_MARGEN_FILTROS = """	VAR __DS0FilterTable = 
		TREATAS({"VENTA BRUTA"}, 'Exl A Maestra de Facturas de Venta'[CATEGORIZACION])

	VAR __DS0FilterTable2 = 
		TREATAS({__ANIO__}, 'Calendario'[Año])

	VAR __DS0FilterTable3 = 
		FILTER(
			KEEPFILTERS(VALUES('Exl A Maestra de Facturas de Venta'[categoria_producto])),
			NOT(
				'Exl A Maestra de Facturas de Venta'[categoria_producto] IN {"CHATARRA",
					"SERVICIOS",
					BLANK()}
			)
		)

	VAR __DS0FilterTable4 = 
		FILTER(
			KEEPFILTERS(VALUES('Exl A Maestra de Facturas de Venta'[producto])),
			NOT(
				'Exl A Maestra de Facturas de Venta'[producto] IN {"PAVO C/M C/ASA EP CONG (8 KG)",
					"ALIMENTACION COMERCIAL"}
			)
		)

	VAR __DS0FilterTable5 = 
		FILTER(
			KEEPFILTERS(VALUES('Exl A Maestra de Facturas de Venta'[estado])),
			NOT('Exl A Maestra de Facturas de Venta'[estado] IN {"Cancelado"})
		)

	VAR __DS0FilterTable6 = 
		FILTER(
			KEEPFILTERS(VALUES('Maestra de Facturacion (Total)'[categoria_producto])),
			NOT(
				'Maestra de Facturacion (Total)'[categoria_producto] IN {BLANK(),
					"BONIFICACION Y REBATES",
					"INTERESES",
					"MATERIA PRIMA",
					"SUMINISTROS",
					"SERVICIOS"}
			)
		)

	VAR __DS0FilterTable7 = 
		FILTER(
			KEEPFILTERS(VALUES('Maestra de Facturacion (Total)'[estado])),
			NOT('Maestra de Facturacion (Total)'[estado] IN {"Cancelado"})
		)

	VAR __DS0FilterTable8 = 
		FILTER(
			KEEPFILTERS(VALUES('Maestra de Kardex (Total)'[CUENTA ORIGEN])),
			NOT('Maestra de Kardex (Total)'[CUENTA ORIGEN] IN {BLANK()})
		)

	VAR __DS0FilterTable9 = 
		FILTER(
			KEEPFILTERS(VALUES('OrdenxFactura'[categoria_producto])),
			NOT(
				'OrdenxFactura'[categoria_producto] IN {BLANK(),
					"BONIFICACION Y REBATES",
					"CHATARRA",
					"SERVICIOS"}
			)
		)

	VAR __DS0FilterTable10 = 
		FILTER(
			KEEPFILTERS(VALUES('OrdenxFactura'[estado])),
			NOT('OrdenxFactura'[estado] IN {"Cancelado"})
		)

	VAR __DS0FilterTable11 = 
		FILTER(
			KEEPFILTERS(VALUES('OrdenxFactura'[producto])),
			NOT(
				'OrdenxFactura'[producto] IN {"PAVO C/M C/ASA EP CONG (8 KG)",
					"ALIMENTACION COMERCIAL"}
			)
		)
"""

_MARGEN_USADOS = "".join(
    "\t\t\t__DS0FilterTable%s,\n" % ("" if i == 1 else i) for i in range(1, 12)
)


def _q_margen(anio, medida_tbl, medida, alias):
    """Consulta del visual "MARGEN VARIABLE PAUNO" con los 11 filtros reales.

    De la consulta original solo se conserva __DS0Core. Lo demás
    (__DS0PrimaryWindowed, __DS0Secondary, NATURALLEFTOUTERJOIN,
    SUBSTITUTEWITHINDEX) es la maquinaria que Power BI usa para pintar el eje
    y la leyenda del gráfico: no cambia los valores, solo los ordena y los
    empareja. Nosotros queremos las cifras, así que evaluamos el core.

    Se mantienen las CINCO columnas de agrupación del original — las tres de
    la tabla de fechas automática más 'Calendario'[Año]/[MES] — porque quitar
    alguna cambiaría el contexto en que se evalúa la medida. Las filas del
    producto cruzado que no corresponden salen en blanco y se descartan al
    leer.
    """
    ld = MARGEN_LOCALDATE
    return (
        "DEFINE\n"
        + _MARGEN_FILTROS.replace("__ANIO__", str(anio))
        + "\n\tVAR __DS0Core = \n"
        "\t\tSUMMARIZECOLUMNS(\n"
        f"\t\t\t'{ld}'[Año],\n"
        f"\t\t\t'{ld}'[Mes],\n"
        f"\t\t\t'{ld}'[NroMes],\n"
        "\t\t\t'Calendario'[MES],\n"
        "\t\t\t'Calendario'[Año],\n"
        + _MARGEN_USADOS
        + f"\t\t\t\"{alias}\", '{medida_tbl}'[{medida}]\n"
        "\t\t)\n\n"
        "EVALUATE\n\t__DS0Core\n\n"
        f"ORDER BY\n\t'{ld}'[Año], '{ld}'[NroMes]"
    )


# Bloque de filtros del visual "PAUNO" de Precio por Kilo Vs Costo por Kilo
# (Copiar consulta, 2026-09-06). Son DOCE, no once: lleva uno que el visual
# de margen no tiene — excluye "BONIFICACION SELL IN" de [DESCRIPCION] — y
# eso corre la numeración de todos los demás (el año pasa de 2 a 3).
#
# Se escribe completo en vez de derivarlo del bloque de margen: en este
# reporte cada visual resultó tener su propio juego de filtros, y armarlo
# por diferencias es justo el tipo de atajo que ya produjo cifras
# equivocadas en Mermas.
_MARGEN_PK_FILTROS = """	VAR __DS0FilterTable = 
		TREATAS({"VENTA BRUTA"}, 'Exl A Maestra de Facturas de Venta'[CATEGORIZACION])

	VAR __DS0FilterTable2 = 
		FILTER(
			KEEPFILTERS(VALUES('Exl A Maestra de Facturas de Venta'[DESCRIPCION])),
			NOT('Exl A Maestra de Facturas de Venta'[DESCRIPCION] IN {"BONIFICACION SELL IN"})
		)

	VAR __DS0FilterTable3 = 
		TREATAS({__ANIO__}, 'Calendario'[Año])

	VAR __DS0FilterTable4 = 
		FILTER(
			KEEPFILTERS(VALUES('Exl A Maestra de Facturas de Venta'[categoria_producto])),
			NOT(
				'Exl A Maestra de Facturas de Venta'[categoria_producto] IN {"CHATARRA",
					"SERVICIOS",
					BLANK()}
			)
		)

	VAR __DS0FilterTable5 = 
		FILTER(
			KEEPFILTERS(VALUES('Exl A Maestra de Facturas de Venta'[producto])),
			NOT(
				'Exl A Maestra de Facturas de Venta'[producto] IN {"PAVO C/M C/ASA EP CONG (8 KG)",
					"ALIMENTACION COMERCIAL"}
			)
		)

	VAR __DS0FilterTable6 = 
		FILTER(
			KEEPFILTERS(VALUES('Exl A Maestra de Facturas de Venta'[estado])),
			NOT('Exl A Maestra de Facturas de Venta'[estado] IN {"Cancelado"})
		)

	VAR __DS0FilterTable7 = 
		FILTER(
			KEEPFILTERS(VALUES('Maestra de Facturacion (Total)'[categoria_producto])),
			NOT(
				'Maestra de Facturacion (Total)'[categoria_producto] IN {BLANK(),
					"BONIFICACION Y REBATES",
					"INTERESES",
					"MATERIA PRIMA",
					"SUMINISTROS",
					"SERVICIOS"}
			)
		)

	VAR __DS0FilterTable8 = 
		FILTER(
			KEEPFILTERS(VALUES('Maestra de Facturacion (Total)'[estado])),
			NOT('Maestra de Facturacion (Total)'[estado] IN {"Cancelado"})
		)

	VAR __DS0FilterTable9 = 
		FILTER(
			KEEPFILTERS(VALUES('Maestra de Kardex (Total)'[CUENTA ORIGEN])),
			NOT('Maestra de Kardex (Total)'[CUENTA ORIGEN] IN {BLANK()})
		)

	VAR __DS0FilterTable10 = 
		FILTER(
			KEEPFILTERS(VALUES('OrdenxFactura'[categoria_producto])),
			NOT(
				'OrdenxFactura'[categoria_producto] IN {BLANK(),
					"BONIFICACION Y REBATES",
					"CHATARRA",
					"SERVICIOS"}
			)
		)

	VAR __DS0FilterTable11 = 
		FILTER(
			KEEPFILTERS(VALUES('OrdenxFactura'[estado])),
			NOT('OrdenxFactura'[estado] IN {"Cancelado"})
		)

	VAR __DS0FilterTable12 = 
		FILTER(
			KEEPFILTERS(VALUES('OrdenxFactura'[producto])),
			NOT(
				'OrdenxFactura'[producto] IN {"PAVO C/M C/ASA EP CONG (8 KG)",
					"ALIMENTACION COMERCIAL"}
			)
		)
"""

_MARGEN_PK_USADOS = "".join(
    "\t\t\t__DS0FilterTable%s,\n" % ("" if i == 1 else i) for i in range(1, 13)
)


def _q_margen_precio_costo(anio):
    """Precio x Kilo y Costo x Kilo mensuales, con los 12 filtros del visual.

    Ambas medidas salen de la misma consulta, igual que en el gráfico. La app
    venía trayendo 'Costo x Kilo' por la sonda genérica, es decir SIN ninguno
    de estos filtros.
    """
    ld = MARGEN_LOCALDATE
    return (
        "DEFINE\n"
        + _MARGEN_PK_FILTROS.replace("__ANIO__", str(anio))
        + "\n\tVAR __DS0Core = \n"
        "\t\tSUMMARIZECOLUMNS(\n"
        f"\t\t\t'{ld}'[Año],\n"
        "\t\t\t'Calendario'[Mes Corto],\n"
        "\t\t\t'Calendario'[MES],\n"
        + _MARGEN_PK_USADOS
        + "\t\t\t\"Precio_x_Kilo\", 'Medidas Julito'[Precio x Kilo],\n"
        "\t\t\t\"Costo_x_Kilo\", 'Medidas Julito'[Costo x Kilo]\n"
        "\t\t)\n\n"
        "EVALUATE\n\t__DS0Core\n\n"
        "ORDER BY\n\t'Calendario'[MES]"
    )


def _q_margen_ventas(anio):
    """Ventas mensuales (S/.) del visual "VENTAS (S/.) PAUNO".

    Copiar consulta 2026-09-06. Comparte los DOCE filtros del visual de
    precio/costo — verificado por diff carácter a carácter contra la
    consulta original, no asumido. Cambian dos cosas: agrupa además por
    'Calendario'[Año], y la cifra es una suma de columna, no una medida:
    SUM('Exl A Maestra de Facturas de Venta'[Monto_Neto_Factura]).

    Como en el visual de margen, se conserva solo __DS0Core; el resto es la
    maquinaria de eje y leyenda.
    """
    ld = MARGEN_LOCALDATE
    return (
        "DEFINE\n"
        + _MARGEN_PK_FILTROS.replace("__ANIO__", str(anio))
        + "\n\tVAR __DS0Core = \n"
        "\t\tSUMMARIZECOLUMNS(\n"
        f"\t\t\t'{ld}'[Año],\n"
        "\t\t\t'Calendario'[Mes Corto],\n"
        "\t\t\t'Calendario'[MES],\n"
        "\t\t\t'Calendario'[Año],\n"
        + _MARGEN_PK_USADOS
        + "\t\t\t\"SumMonto_Neto_Factura\", "
        "CALCULATE(SUM('Exl A Maestra de Facturas de Venta'[Monto_Neto_Factura]))\n"
        "\t\t)\n\n"
        "EVALUATE\n\t__DS0Core\n\n"
        "ORDER BY\n\t'Calendario'[MES]"
    )


def serie_margen_ventas(token, ds_id, periodos):
    """Serie mensual de Ventas (S/.)."""
    anios = sorted({a for a, _ in periodos})
    mapa = {}
    for anio in anios:
        for r in dax(token, ds_id, _q_margen_ventas(anio), f"margen-ventas-{anio}") or []:
            v = r.get("[SumMonto_Neto_Factura]", r.get("SumMonto_Neto_Factura"))
            if v is None:
                continue
            mes = r.get("Calendario[MES]") or r.get("[MES]")
            if isinstance(mes, str):
                mes = MESES_CORTOS.get(mes.strip().lower()[:3])
            if mes is None:
                continue
            try:
                mapa[(int(anio), int(mes))] = float(v)
            except (TypeError, ValueError):
                pass
    serie = [mapa.get(pp) for pp in periodos]
    if not any(x is not None for x in serie):
        print("    ✗ Ventas (S/.): sin datos")
        return {}
    print(f"    [Ventas (S/.)]: "
          f"{sum(1 for x in serie if x is not None)}/{len(serie)} meses")
    return {"Ventas (S/.)": serie}


def serie_margen_precio_costo(token, ds_id, periodos):
    """Series mensuales de Precio x Kilo y Costo x Kilo."""
    anios = sorted({a for a, _ in periodos})
    pares = {"Precio x Kilo": {}, "Costo x Kilo": {}}
    for anio in anios:
        for r in dax(token, ds_id, _q_margen_precio_costo(anio),
                     f"margen-pk-{anio}") or []:
            mes = r.get("Calendario[MES]") or r.get("[MES]")
            if isinstance(mes, str):
                mes = MESES_CORTOS.get(mes.strip().lower()[:3])
            if mes is None:
                continue
            for alias, etiqueta in (("Precio_x_Kilo", "Precio x Kilo"),
                                    ("Costo_x_Kilo", "Costo x Kilo")):
                v = r.get(f"[{alias}]", r.get(alias))
                if v is None:
                    continue
                try:
                    pares[etiqueta][(int(anio), int(mes))] = float(v)
                except (TypeError, ValueError):
                    pass

    out = {}
    for etiqueta, mapa in pares.items():
        serie = [mapa.get(pp) for pp in periodos]
        if any(x is not None for x in serie):
            out[etiqueta] = serie
            print(f"    [{etiqueta}]: "
                  f"{sum(1 for x in serie if x is not None)}/{len(serie)} meses")
    return out


def serie_margen(token, ds_id, periodos, medidas):
    """Series mensuales del reporte de Margen.

    `medidas` es una lista de (etiqueta, tabla, medida).
    """
    anios = sorted({a for a, _ in periodos})
    ld = MARGEN_LOCALDATE
    out = {}
    for etiqueta, tbl, medida in medidas:
        alias = "v_" + re.sub(r"[^A-Za-z0-9]", "_", etiqueta)
        por_periodo = {}
        for anio in anios:
            q = _q_margen(anio, tbl, medida, alias)
            for r in dax(token, ds_id, q, f"margen-{etiqueta}-{anio}") or []:
                val = r.get(f"[{alias}]", r.get(alias))
                if val is None:
                    continue          # fila del producto cruzado sin dato
                mes = r.get("Calendario[MES]") or r.get(f"{ld}[NroMes]")
                if isinstance(mes, str):
                    mes = MESES_CORTOS.get(mes.strip().lower()[:3])
                if mes is None:
                    continue
                try:
                    por_periodo[(int(anio), int(mes))] = float(val)
                except (TypeError, ValueError):
                    pass
        serie = [por_periodo.get(pp) for pp in periodos]
        if any(x is not None for x in serie):
            out[etiqueta] = serie
            print(f"    [{etiqueta}]: "
                  f"{sum(1 for x in serie if x is not None)}/{len(serie)} meses")
        else:
            print(f"    ✗ {etiqueta}: sin datos ({medida})")
    return out


def serie_mermas_exacta(token, ds_id, periodos):
    """Series mensuales de Mermas con los filtros reales del visual.

    Devuelve {nombre_medida: [valores por período]} o None.
    Se consulta un año por vez porque el visual filtra por el segmentador ANO.
    """
    por_periodo = {alias: {} for _, alias in MERMAS_MEDIDAS}
    anios = sorted({a for a, _ in periodos})
    for anio in anios:
        rows = dax(token, ds_id, _q_mermas(anio), f"mermas-{anio}")
        for r in rows or []:
            mes = r.get("Calendario[MES]") or r.get("[MES]")
            if mes is None:
                continue
            if isinstance(mes, str):
                mes = MESES_CORTOS.get(mes.strip().lower()[:3])
            if mes is None:
                continue
            for col, alias in MERMAS_MEDIDAS:
                v = r.get(f"[{col}]", r.get(col))
                if v is None:
                    continue
                try:
                    por_periodo[alias][(int(anio), int(mes))] = float(v)
                except (TypeError, ValueError):
                    pass

    out = {}
    for _, alias in MERMAS_MEDIDAS:
        serie = [por_periodo[alias].get(p) for p in periodos]
        if any(v is not None for v in serie):
            out[alias] = serie
    return out or None


def serie_mermas_plantas(token, ds_id, periodos):
    """Series mensuales de % Merma por planta.

    Solo las de PLANTAS_MERMAS: cada una con su almacén y su medida, ambos
    confirmados por Copiar consulta. Ver la nota de esa constante.
    """
    anios = sorted({a for a, _ in periodos})
    out = {}
    for etiqueta, almacenes, medida, tipo_base in PLANTAS_MERMAS:
        alias = "v_" + re.sub(r"[^A-Za-z0-9]", "_", etiqueta)
        por_periodo = {}
        for anio in anios:
            q = _q_mermas_planta(anio, almacenes, medida, alias, tipo_base)
            for r in dax(token, ds_id, q, f"mermas-{etiqueta}-{anio}") or []:
                mes = r.get("Calendario[MES]") or r.get("[MES]")
                if isinstance(mes, str):
                    mes = MESES_CORTOS.get(mes.strip().lower()[:3])
                val = r.get(f"[{alias}]", r.get(alias))
                if mes is None or val is None:
                    continue
                try:
                    por_periodo[(int(anio), int(mes))] = float(val)
                except (TypeError, ValueError):
                    pass
        serie = [por_periodo.get(pp) for pp in periodos]
        if any(x is not None for x in serie):
            out[f"% Merma {etiqueta}"] = serie
            print(f"    [% Merma {etiqueta}]: "
                  f"{sum(1 for x in serie if x is not None)}/{len(serie)} meses "
                  f"({'+'.join(almacenes)} / {medida})")
        else:
            print(f"    ✗ % Merma {etiqueta}: sin datos")
    return out


def serie_mermas_segmentos(token, ds_id, periodos):
    """Series mensuales de % Merma por unidad de negocio.

    Solo los segmentos de SEGMENTOS_MERMAS: cada uno con la medida y el juego
    de filtros que realmente usa su visual. Ver la nota de esa constante sobre
    por qué se retiró el sondeo por nombre.
    """
    anios = sorted({a for a, _ in periodos})
    out = {}
    for etiqueta, medida, tipo_base in SEGMENTOS_MERMAS:
        alias = "v_" + re.sub(r"[^A-Za-z0-9]", "_", etiqueta)
        por_periodo = {}
        for anio in anios:
            q = _q_mermas_segmento(anio, "Tabla Mermas", medida, alias, tipo_base)
            for r in dax(token, ds_id, q, f"mermas-{etiqueta}-{anio}") or []:
                mes = r.get("Calendario[MES]") or r.get("[MES]")
                if isinstance(mes, str):
                    mes = MESES_CORTOS.get(mes.strip().lower()[:3])
                val = r.get(f"[{alias}]", r.get(alias))
                if mes is None or val is None:
                    continue
                try:
                    por_periodo[(int(anio), int(mes))] = float(val)
                except (TypeError, ValueError):
                    pass
        serie = [por_periodo.get(pp) for pp in periodos]
        if any(x is not None for x in serie):
            out[f"% Merma {etiqueta}"] = serie
            print(f"    [% Merma {etiqueta}]: "
                  f"{sum(1 for x in serie if x is not None)}/{len(serie)} meses "
                  f"({medida})")
        else:
            print(f"    ✗ % Merma {etiqueta}: sin datos ({medida})")
    return out


def main():
    print("=== JUANITO — SERIE HISTÓRICA MENSUAL ===\n")
    token = get_token()
    print("Token OK\n")

    cache_file = OUTPUT_DIR / "measure_cache.json"
    if not cache_file.exists():
        print("ERROR: falta measure_cache.json — corre fetch_powerbi.py primero.")
        return
    cache = json.loads(cache_file.read_text())

    periodos = build_periodos()
    periodos_str = [f"{a}-{m:02d}" for a, m in periodos]
    print(f"Períodos: {periodos_str[0]} → {periodos_str[-1]} ({len(periodos)} meses)\n")

    resultado = {}
    fechas_usadas = {}
    sin_serie = {}

    # ── CxC: consulta exacta del gráfico (la sonda genérica no la alcanza)
    cxc_id = DATASET_IDS.get("cxc")
    if cxc_id:
        print("── cxc (consulta exacta del gráfico '% DE MOROSIDAD')")
        try:
            s = serie_cxc_morosidad(token, cxc_id, periodos)
            if s:
                con_dato = sum(1 for v in s if v is not None)
                resultado.setdefault("cxc", {})["% Morosidad"] = s
                fechas_usadas["cxc"] = f"{CXC_DATE_TBL}[Año] + 'Calendario (FACT)'[Mes Numero]"
                print(f"    [% Morosidad]: {con_dato}/{len(s)} meses con dato — "
                      f"medida '[% MOROSIDAD x mes]'\n")
            else:
                print("    Sin filas — la consulta no devolvió datos\n")
        except Exception as e:
            print(f"    ✗ {e}\n")

    # ── Mermas: consulta exacta del gráfico, con los 6 filtros del visual.
    # La sonda genérica sí devolvía una serie, pero SIN esos filtros: incluía
    # chatarra, materia prima, servicios y comprobantes cancelados, así que
    # el % no coincidía con el que ve el equipo en el reporte.
    mermas_id = DATASET_IDS.get("mermas")
    if mermas_id:
        print("── mermas (consulta exacta del gráfico 'MERMA MENSUAL POR PLANTA')")
        try:
            s = serie_mermas_exacta(token, mermas_id, periodos)
            if s:
                resultado["mermas"] = s
                fechas_usadas["mermas"] = "Calendario[Año] + Calendario[MES] (6 filtros del visual)"
                for med, vals in s.items():
                    print(f"    [{med}]: {sum(1 for v in vals if v is not None)}/{len(vals)} meses")
                # Segmentos (B&D, TIGO, MAQUILA, plantas...): mismo bloque de
                # filtros menos el de TIPO DE BASE, con la medida por segmento.
                resultado["mermas"].update(serie_mermas_segmentos(token, mermas_id, periodos))
                # Plantas: la medida base filtrada por 'Tabla Mermas'[almacen]
                resultado["mermas"].update(serie_mermas_plantas(token, mermas_id, periodos))
                print()
            else:
                print("    Sin filas\n")
        except Exception as e:
            print(f"    ✗ {e}\n")

    # ── Margen: consulta exacta del visual, con sus 11 filtros
    margen_id = DATASET_IDS.get("margen")
    if margen_id:
        print("── margen (consulta exacta de 'MARGEN VARIABLE PAUNO')")
        try:
            s = serie_margen(token, margen_id, periodos, [
                ("% Margen Variable", "Medidas Julito", "% Margen Contribución. %"),
            ])
            if s:
                resultado.setdefault("margen", {}).update(s)
            # Precio y costo por kilo: otro visual, otros 12 filtros.
            # Sustituyen al 'Costo x Kilo' de la sonda genérica, que venía
            # sin ningún filtro de negocio.
            resultado.setdefault("margen", {}).update(
                serie_margen_precio_costo(token, margen_id, periodos))
            # Ventas: mismos 12 filtros (verificado por diff), otra columna
            resultado["margen"].update(
                serie_margen_ventas(token, margen_id, periodos))
            print()
        except Exception as e:
            print(f"    ✗ {e}\n")

    for ds_key, medidas_dict in cache.items():
        ds_id = DATASET_IDS.get(ds_key)
        if not ds_id:
            continue
        if ds_key == "mermas" and "mermas" in resultado:
            # ya resuelto con la consulta exacta; la genérica solo añadiría
            # una serie sin los filtros del visual (números no comparables)
            continue
        if ds_key == "cxc" and "cxc" in resultado:
            # ya resuelto arriba con la consulta exacta; la sonda genérica
            # solo volvería a marcar '% Morosidad' como plana
            continue
        medidas = list(medidas_dict.keys())
        if not medidas:
            continue

        print(f"── {ds_key} ({len(medidas)} medidas)")
        tbl, col = detectar_fecha(token, ds_id, medidas[0])
        if not tbl:
            # Deja constancia de QUÉ contiene el modelo, para no volver a adivinar
            print("    Sin tabla de fecha que filtre las medidas — diagnosticando el modelo...")
            diag = {"medidas": medidas, "tablas": [], "columnas_fecha": []}
            t_rows = dax(token, ds_id, "EVALUATE INFO.TABLES()", "diag-tables", retries=1)
            if t_rows:
                diag["tablas"] = sorted({(r.get("[Name]") or r.get("Name") or "") for r in t_rows} - {""})
            diag["columnas_fecha"] = [f"{t}[{c}]" for t, c in columnas_fecha_reales(token, ds_id)]
            if not t_rows:
                # DMV bloqueada: cae al esquema REST
                r = requests.get(f"{PBI_BASE}/groups/{WS_ID}/datasets/{ds_id}/tables",
                                 headers={"Authorization": f"Bearer {token}"}, timeout=25)
                if r.ok:
                    diag["tablas"] = [t.get("name") for t in r.json().get("value", [])]
                    diag["esquema_rest"] = {
                        t.get("name"): [c.get("name") for c in t.get("columns", [])]
                        for t in r.json().get("value", [])
                    }
                else:
                    diag["esquema_rest_error"] = f"HTTP {r.status_code}"
            sin_serie[ds_key] = diag
            print(f"    Tablas del modelo: {diag['tablas'][:15]}")
            print(f"    Columnas de fecha: {diag['columnas_fecha'][:15]}")
            print(f"    → estas medidas no admiten corte mensual con el modelo actual\n")
            continue
        fechas_usadas[ds_key] = f"{tbl}[{col}]"

        serie = serie_batch(token, ds_id, medidas, tbl, col, periodos)
        if serie is None:
            print("    Batch falló — usando mes a mes")
            serie = serie_mes_a_mes(token, ds_id, medidas, tbl, col, periodos)

        # Reporta cobertura y detecta series "planas": la medida no responde
        # al filtro de fecha genérico (CALCULATE + FILTER(ALL(Calendario)))
        # y Power BI devuelve el mismo total sin importar el mes — mismo bug
        # confirmado a mano en Control Interno/CxP/Inventario/Avance/Compras.
        # No se puede saber de antemano qué medida tiene este problema sin
        # Copiar consulta, así que se detecta por evidencia: casi todos los
        # valores no nulos son idénticos.
        serie_confiable = {}
        for med, vals in serie.items():
            con_dato = sum(1 for v in vals if v is not None)
            no_nulos = [v for v in vals if v is not None]
            # Misma regla que usa el dashboard en el navegador (3 pruebas):
            media = (sum(no_nulos) / len(no_nulos)) if no_nulos else 0
            # 1) pocos valores distintos (redondeo RELATIVO: el ruido de coma
            #    flotante no debe partir un mismo valor en dos buckets)
            distintos = len({round(v / media * 1e4) for v in no_nulos}) if media else 0
            pocas_variantes = con_dato >= 3 and distintos <= max(2, con_dato // 7)
            # 2) amplitud relativa < 1% => ruido, no tendencia
            amplitud = (max(no_nulos) - min(no_nulos)) / media if (no_nulos and media) else 0
            plana_por_amplitud = con_dato >= 3 and abs(amplitud) < 0.01
            # 3) valor congelado: una misma cifra ocupa más de la mitad de los meses
            moda = 0
            if no_nulos and media:
                conteo = {}
                for v in no_nulos:
                    k = round(v / media * 1e4)
                    conteo[k] = conteo.get(k, 0) + 1
                moda = max(conteo.values())
            congelada = con_dato >= 3 and moda > len(no_nulos) / 2
            es_plana = pocas_variantes or plana_por_amplitud or congelada
            estado = "PLANA — no confiable" if es_plana else "OK"
            print(f"    [{med}]: {con_dato}/{len(vals)} meses con dato, {distintos} valores distintos → {estado}")
            if es_plana:
                sin_serie.setdefault(ds_key, {"medidas": [], "tablas": [tbl], "columnas_fecha": [f"{tbl}[{col}]"]})
                sin_serie[ds_key].setdefault("medidas_planas", []).append({
                    "medida": med, "valores_distintos": distintos, "meses_con_dato": con_dato,
                    "amplitud_relativa_pct": round(abs(amplitud) * 100, 4),
                    "repeticiones_del_valor_mas_comun": moda,
                    "motivo": "La medida no varía por mes con el filtro CALCULATE+FILTER(ALL(Calendario)) — "
                              "posible relación inactiva o la medida ignora el filtro de fecha (confirmado con "
                              "Copiar consulta solo en Control Interno; para las demás falta validar en vivo).",
                })
            else:
                serie_confiable[med] = vals

        if serie_confiable:
            # update, no asignación: 'margen' ya trae la serie de la consulta
            # exacta y asignar la pisaría. Las capturadas mandan, así que las
            # de la sonda genérica solo rellenan lo que falta.
            base = resultado.setdefault(ds_key, {})
            for med, vals in serie_confiable.items():
                base.setdefault(med, vals)
        print()

    out = {
        "generado": datetime.date.today().isoformat(),
        "fuente": "Power BI REST API executeQueries",
        "periodos": periodos_str,
        "tablas_fecha": fechas_usadas,
        "sin_serie": sin_serie,
        "datasets": resultado,
    }
    path = OUTPUT_DIR / "series.json"
    # json.dumps escribe Infinity/NaN, que NO son JSON válido: el navegador
    # falla en JSON.parse y la app se queda sin NINGUNA serie. Aparecen de
    # verdad — p.ej. '% Merma <almacén>' en meses sin producción divide entre
    # cero — así que se convierten a null (mes sin dato) en vez de ocultarlos.
    # allow_nan=False haría reventar el script; preferimos publicar el resto.
    n_inf = _sanear_no_finitos(out)
    if n_inf:
        print(f"⚠ {n_inf} valores no finitos (Infinity/NaN) → null")
    # Series que quedan con menos de 3 meses útiles no sirven para tendencia
    for ds_key in list(out["datasets"]):
        for med in list(out["datasets"][ds_key]):
            vals = out["datasets"][ds_key][med]
            if sum(1 for v in vals if v is not None) < 3:
                del out["datasets"][ds_key][med]
                print(f"⚠ [{ds_key}] {med}: <3 meses con dato — se descarta")
        if not out["datasets"][ds_key]:
            del out["datasets"][ds_key]
    path.write_text(json.dumps(out, ensure_ascii=False, indent=2, allow_nan=False))
    print(f"Guardado: {path}")
    print(f"Datasets con serie: {list(resultado.keys())}")


if __name__ == "__main__":
    main()
