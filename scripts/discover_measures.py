#!/usr/bin/env python3
"""Descubre tabla de fechas en cada dataset y prueba medidas con filtro de mes actual."""
import os, requests
from datetime import date

TENANT_ID     = os.environ["AZURE_TENANT_ID"]
CLIENT_ID     = os.environ["AZURE_CLIENT_ID"]
CLIENT_SECRET = os.environ["AZURE_CLIENT_SECRET"]
USERNAME      = os.environ["PBI_USERNAME"]
PASSWORD      = os.environ["PBI_PASSWORD"]
WS_ID         = "461932ad-b5ec-4fd6-aa97-f1fc7bdc5169"

TOKEN_URL = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
PBI_BASE  = "https://api.powerbi.com/v1.0/myorg"

r = requests.post(TOKEN_URL, data={
    "grant_type": "password", "client_id": CLIENT_ID,
    "client_secret": CLIENT_SECRET, "username": USERNAME,
    "password": PASSWORD, "scope": "https://analysis.windows.net/powerbi/api/.default",
})
token = r.json()["access_token"]
H  = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

today = date.today()
YEAR  = today.year
MONTH = today.month

def dax_raw(ds_id, query):
    url = f"{PBI_BASE}/groups/{WS_ID}/datasets/{ds_id}/executeQueries"
    resp = requests.post(url, headers=H,
                         json={"queries": [{"query": query}],
                               "serializerSettings": {"includeNulls": True}},
                         timeout=20)
    if not resp.ok:
        return None, resp.text[:200]
    rows = resp.json().get("results",[{}])[0].get("tables",[{}])[0].get("rows",[])
    return rows, None

# Datasets clave a probar
DATASETS = {
    "margen":    "38076daa-d2cd-4a93-858a-82c0a4cf8cb6",
    "fill_rate": "7f4ebe22-5e90-4e35-973b-4af3c58497e5",
    "compras":   "06408938-8202-424e-80c0-b42c178dabde",
    "cxc":       "2eec70cd-0820-408f-b938-a2cd547b0c18",
    "mermas":    "35866214-f4da-45a3-a5a2-aa0c8caffe78",
}

# Medidas a probar por dataset
MEASURES = {
    "margen":    "Venta Total",
    "fill_rate": "% FILLRATE",
    "compras":   None,  # usa SUM de columna
    "cxc":       "% Morosidad",
    "mermas":    "% Merma Total",
}

# Patrones de tablas de fecha comunes en Power BI peruano
DATE_TABLES = [
    ("Calendario", "Fecha"),
    ("Calendario", "Date"),
    ("Calendar",   "Date"),
    ("Calendar",   "Fecha"),
    ("Fechas",     "Fecha"),
    ("Fechas",     "Date"),
    ("DimFecha",   "Fecha"),
    ("DimDate",    "Date"),
    ("Tiempo",     "Fecha"),
    ("Dim_Fecha",  "Fecha"),
    ("Date",       "Date"),
    ("Fecha",      "Fecha"),
]

def probe_date_table(ds_id, tbl, col):
    """Prueba si la tabla/columna de fecha existe y tiene datos del mes actual."""
    q = f"""EVALUATE
CALCULATETABLE(
  ROW("n", COUNTROWS('{tbl}')),
  FILTER(ALL('{tbl}'), YEAR('{tbl}'[{col}]) = {YEAR} && MONTH('{tbl}'[{col}]) = {MONTH})
)"""
    rows, err = dax_raw(ds_id, q)
    if rows and rows[0]:
        v = list(rows[0].values())[0]
        return v
    return None

def measure_with_date(ds_id, measure_name, tbl, col):
    """Ejecuta medida con filtro de mes actual."""
    q = f"""EVALUATE
ROW("v",
  CALCULATE(
    [{measure_name}],
    FILTER(ALL('{tbl}'), YEAR('{tbl}'[{col}]) = {YEAR} && MONTH('{tbl}'[{col}]) = {MONTH})
  )
)"""
    rows, err = dax_raw(ds_id, q)
    if rows and rows[0]:
        return list(rows[0].values())[0]
    return None

def measure_no_filter(ds_id, measure_name):
    """Medida sin filtro (valor actual)."""
    rows, err = dax_raw(ds_id, f'EVALUATE ROW("v", [{measure_name}])')
    if rows and rows[0]:
        return list(rows[0].values())[0]
    return None

print(f"=== DESCUBRIMIENTO TABLA FECHA — {YEAR}-{MONTH:02d} ===\n")

for ds_key, ds_id in DATASETS.items():
    measure = MEASURES.get(ds_key)
    print(f"\n--- {ds_key.upper()} ({ds_id[:8]}) ---")

    # Sin filtro (valor actual — erróneo)
    if measure:
        v_total = measure_no_filter(ds_id, measure)
        print(f"  Sin filtro [{measure}] = {v_total}")

    # Buscar tabla de fecha
    found_tbl, found_col, found_rows = None, None, None
    for tbl, col in DATE_TABLES:
        n = probe_date_table(ds_id, tbl, col)
        if n is not None:
            print(f"  ✓ Tabla fecha: '{tbl}'[{col}] — {n} filas en {YEAR}-{MONTH:02d}")
            found_tbl, found_col, found_rows = tbl, col, n
            break
        else:
            print(f"    ✗ '{tbl}'[{col}]")

    # Aplicar filtro de mes si encontramos la tabla
    if found_tbl and measure:
        v_mes = measure_with_date(ds_id, measure, found_tbl, found_col)
        print(f"  Con filtro mes [{measure}] = {v_mes}")

        # También probar mes anterior
        prev_month = MONTH - 1 if MONTH > 1 else 12
        prev_year  = YEAR if MONTH > 1 else YEAR - 1
        q_prev = f"""EVALUATE
ROW("v",
  CALCULATE(
    [{measure}],
    FILTER(ALL('{found_tbl}'), YEAR('{found_tbl}'[{found_col}]) = {prev_year} && MONTH('{found_tbl}'[{found_col}]) = {prev_month})
  )
)"""
        rows, _ = dax_raw(ds_id, q_prev)
        v_prev = list(rows[0].values())[0] if rows and rows[0] else None
        print(f"  Mes anterior [{measure}] = {v_prev}")

    # Para fill_rate, probar también las otras medidas
    if ds_key == "fill_rate" and found_tbl:
        for m in ["ORDEN DE VENTA", "FACTURACION", "VENTA PERDIDA"]:
            v = measure_with_date(ds_id, m, found_tbl, found_col)
            print(f"  Con filtro mes [{m}] = {v}")

print("\n=== FIN ===")
