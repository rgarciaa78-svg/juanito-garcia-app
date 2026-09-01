#!/usr/bin/env python3
"""Obtiene IDs completos de todos los datasets y prueba más medidas de Fill Rate y Productividad."""
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
H  = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
HG = {"Authorization": f"Bearer {token}"}

def try_m(ds_id, name):
    url = f"{PBI_BASE}/groups/{WS_ID}/datasets/{ds_id}/executeQueries"
    resp = requests.post(url, headers=H,
                         json={"queries": [{"query": f'EVALUATE ROW("v", [{name}])'}],
                               "serializerSettings": {"includeNulls": True}},
                         timeout=15)
    if resp.ok:
        rows = resp.json().get("results",[{}])[0].get("tables",[{}])[0].get("rows",[])
        v = list(rows[0].values())[0] if rows else None
        return v, True
    return None, False

# 1. Imprimir TODOS los datasets con ID completo
print("=== TODOS LOS DATASETS (ID COMPLETO) ===")
ds_r = requests.get(f"{PBI_BASE}/groups/{WS_ID}/datasets", headers=HG)
all_datasets = ds_r.json().get("value", [])
calculo_id = None
for ds in all_datasets:
    print(f"  {ds['id']}  {ds['name']}")
    if "alculo" in ds["name"] or "rovision" in ds["name"] or "7f4ebe22" in ds["id"]:
        calculo_id = ds["id"]
        print(f"  ^^^ FILL RATE DATASET: {calculo_id}")

# 2. Probar más medidas en el dataset de Calculo de Provisiones
if calculo_id:
    print(f"\n=== MEDIDAS ADICIONALES EN CALCULO PROVISIONES ({calculo_id[:8]}) ===")
    extra = [
        "% FILLRATE", "% Fill Rate", "% FillRate",
        "ORDEN DE VENTA", "FACTURACION", "VENTA PERDIDA", "AMONESTACION",
        "Fill Rate Mes", "% Fill Rate Mes", "FR Mes",
        "Venta Perdida", "Orden de Venta", "Facturacion",
        "Amonestacion", "Amonestación",
    ]
    for name in extra:
        v, ok = try_m(calculo_id, name)
        if ok:
            print(f"  ✓ [{name}] = {v}")

# 3. Más candidatos Fill Rate en margen (ya confirmado que tiene las medidas)
MARGEN_ID = "38076daa-d2cd-4a93-858a-82c0a4cf8cb6"
print(f"\n=== FILL RATE EN MARGEN — mas variantes ===")
fr_more = [
    "% FILLRATE", "% Fill Rate", "% FillRate",
    "ORDEN DE VENTA", "FACTURACION", "VENTA PERDIDA", "AMONESTACION",
    "Fill Rate Mes", "% Fill Rate Mes",
    "Fill Rate Acumulado", "% Fill Rate Acum",
    "Venta Perdida", "Orden de Venta",
]
for name in fr_more:
    v, ok = try_m(MARGEN_ID, name)
    if ok:
        print(f"  ✓ [{name}] = {v}")

print("\n=== FIN ===")
