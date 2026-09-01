#!/usr/bin/env python3
"""Prueba medidas de Productividad con y sin filtro de fecha."""
import os, requests

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
H = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

def dax_raw(ds_id, q):
    url = f"{PBI_BASE}/groups/{WS_ID}/datasets/{ds_id}/executeQueries"
    resp = requests.post(url, headers=H,
                         json={"queries": [{"query": q}],
                               "serializerSettings": {"includeNulls": True}},
                         timeout=20)
    if resp.ok:
        rows = resp.json().get("results",[{}])[0].get("tables",[{}])[0].get("rows",[])
        return rows, None
    err = resp.json().get("error",{}).get("pbi.error",{}).get("code","?")
    return [], err

PROD_ID = "a042ba6a-c82c-4fc1-bf95-b9e84bd15fc6"

# 1. Descubrir columnas del Calendario en Productividad
print("=== CALENDARIO PRODUCTIVIDAD ===")
for q in [
    "EVALUATE TOPN(3, 'Calendario')",
    "EVALUATE TOPN(3, 'Calendar')",
    "EVALUATE TOPN(3, 'Fecha')",
]:
    rows, err = dax_raw(PROD_ID, q)
    if rows:
        print(f"  ✓ {q[:30]}: {list(rows[0].keys())[:6]}")
        for row in rows[:2]: print(f"    {row}")
    else:
        print(f"  ✗ {err}")

# 2. Probar medidas con filtro de año=2026
print("\n=== MEDIDAS CON FILTRO ANO=2026 ===")
NAMES = [
    "Produccion Total (KG)", "PRODUCCION TOTAL (KG)", "Produccion Total",
    "Producción Total (KG)", "Producción Total",
    "KG Total", "Total KG", "KG Produccion", "KG Prod",
    "Venta Neta (KG)", "VENTA NETA (KG)", "Venta Neta",
    "Planilla Total (S/.)", "PLANILLA TOTAL (S/.)", "Planilla Total",
    "Planilla (S/.) entre Kg Producido", "PLANILLA (S/.) ENTRE KG PRODUCIDO",
    "Planilla entre KG", "S/Kg Producido", "Planilla KG",
    "Planilla", "Total Planilla", "Costo Planilla",
    "MOD", "Costo MOD", "Gasto MOD",
    "Kg Producidos", "Kg Producción", "Kg Prod",
    "Eficiencia", "Productividad", "Ratio MOD",
]

found = {}
for name in NAMES:
    # Sin filtro
    rows, err = dax_raw(PROD_ID, f'EVALUATE ROW("v", [{name}])')
    if rows:
        v = list(rows[0].values())[0]
        found[name] = v
        print(f"  ✓ (sin filtro) [{name}] = {v}")
        continue
    # Con filtro Calendario[Año] = 2026
    rows2, err2 = dax_raw(PROD_ID,
        f'EVALUATE CALCULATETABLE(ROW("v", [{name}]), \'Calendario\'[Año] = 2026)')
    if rows2:
        v = list(rows2[0].values())[0]
        found[name] = v
        print(f"  ✓ (con filtro) [{name}] = {v}")
    else:
        print(f"  ✗ [{name}] — {err}/{err2}")

print(f"\nEncontradas: {found}")

# 3. Intentar SUMMARIZE sobre Medidas para ver qué columnas existen
print("\n=== MEDIDAS tabla ===")
for q in [
    "EVALUATE 'Medidas'",
    "EVALUATE VALUES('Medidas'[Column])",
    "EVALUATE TOPN(20, 'Medidas')",
    "EVALUATE SELECTCOLUMNS('Medidas', \"col\", 'Medidas'[Column])",
]:
    rows, err = dax_raw(PROD_ID, q)
    if rows:
        print(f"  ✓ {q[:40]}: {rows[:5]}")
    else:
        print(f"  ✗ {q[:40]}: {err}")

print("\n=== FIN ===")
