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

import os, json, time, datetime, requests
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


def detectar_fecha(token, dataset_id, medida):
    """Prueba pares (tabla, columna) hasta que uno responda con un filtro de mes."""
    for tbl, col in DATE_CANDIDATES:
        q = (f'EVALUATE ROW("v", CALCULATE([{medida}], '
             f"FILTER(ALL('{tbl}'), YEAR('{tbl}'[{col}]) = 2025 && MONTH('{tbl}'[{col}]) = 6)))")
        rows = dax(token, dataset_id, q, f"probe {tbl}[{col}]", retries=1)
        if rows is not None and len(rows) > 0:
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

    for ds_key, medidas_dict in cache.items():
        ds_id = DATASET_IDS.get(ds_key)
        if not ds_id:
            continue
        medidas = list(medidas_dict.keys())
        if not medidas:
            continue

        print(f"── {ds_key} ({len(medidas)} medidas)")
        tbl, col = detectar_fecha(token, ds_id, medidas[0])
        if not tbl:
            print(f"    Sin tabla de fecha reconocible — se omite la serie\n")
            continue
        fechas_usadas[ds_key] = f"{tbl}[{col}]"

        serie = serie_batch(token, ds_id, medidas, tbl, col, periodos)
        if serie is None:
            print("    Batch falló — usando mes a mes")
            serie = serie_mes_a_mes(token, ds_id, medidas, tbl, col, periodos)

        # Reporta cobertura
        for med, vals in serie.items():
            con_dato = sum(1 for v in vals if v is not None)
            print(f"    [{med}]: {con_dato}/{len(vals)} meses con dato")

        resultado[ds_key] = serie
        print()

    out = {
        "generado": datetime.date.today().isoformat(),
        "fuente": "Power BI REST API executeQueries",
        "periodos": periodos_str,
        "tablas_fecha": fechas_usadas,
        "datasets": resultado,
    }
    path = OUTPUT_DIR / "series.json"
    path.write_text(json.dumps(out, ensure_ascii=False, indent=2))
    print(f"Guardado: {path}")
    print(f"Datasets con serie: {list(resultado.keys())}")


if __name__ == "__main__":
    main()
