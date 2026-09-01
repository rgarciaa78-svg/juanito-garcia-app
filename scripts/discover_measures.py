#!/usr/bin/env python3
"""Busca dataset de Fill Rate y prueba % FILLRATE en todos los datasets del workspace."""
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

# 1. Listar TODOS los datasets del workspace
print("=== TODOS LOS DATASETS ===")
ds_r = requests.get(f"{PBI_BASE}/groups/{WS_ID}/datasets", headers=HG)
all_datasets = ds_r.json().get("value", [])
for ds in all_datasets:
    print(f"  [{ds['id'][:8]}] {ds['name']}")

# 2. Probar % FILLRATE en TODOS los datasets
print("\n=== % FILLRATE EN TODOS LOS DATASETS ===")
fr_candidates = [
    "% FILLRATE", "% Fill Rate", "% FillRate", "FILLRATE",
    "% FR", "Fill Rate", "FR",
    "ORDEN DE VENTA", "FACTURACION", "VENTA PERDIDA", "AMONESTACION",
]

known_datasets = {
    "cxc":            "2eec70cd-0820-408f-b938-a2cd547b0c18",
    "cxp":            "45a8ab8d-e162-4398-a260-3a9a5f90829f",
    "margen":         "38076daa-d2cd-4a93-858a-82c0a4cf8cb6",
    "mermas":         "35866214-f4da-45a3-a5a2-aa0c8caffe78",
    "compras":        "06408938-8202-424e-80c0-b42c178dabde",
    "inventario":     "0e27d784-41a4-48f0-9208-60210119f0a7",
    "control_ds":     "0aca7bdd-6b72-49c2-be41-17ae0f6b5848",
    "productividad":  "a042ba6a-c82c-4fc1-bf95-b9e84bd15fc6",
    "planificacion":  "30074d92-7ec1-4762-82f2-1cb29c15dcfe",
    "consumo":        "c972c8cb-e5fc-4b60-8f5e-265a78e1e796",
}

# Agregar datasets desconocidos
all_ds_ids = {ds["name"]: ds["id"] for ds in all_datasets}
for ds_name, ds_id in all_ds_ids.items():
    if ds_id not in known_datasets.values():
        known_datasets[ds_name[:20]] = ds_id

for ds_key, ds_id in known_datasets.items():
    for name in fr_candidates:
        v, ok = try_m(ds_id, name)
        if ok:
            print(f"  ✓ [{ds_key}] [{name}] = {v}")

# 3. Para Productividad: probar SUM directo sobre tabla con datos reales
PROD_ID = "a042ba6a-c82c-4fc1-bf95-b9e84bd15fc6"
print("\n=== PRODUCTIVIDAD: tablas y SUM directos ===")
# Intentar nombres de tablas más específicos
prod_tables = [
    "Planilla", "fPlanilla", "hPlanilla",
    "Produccion", "fProduccion", "hProduccion",
    "MOD", "Funcionario", "Funcionarios",
    "Empleado", "Personal", "RR HH", "RRHH",
    "Hechos Planilla", "Hechos Produccion",
    "HechosProductividad", "FactProductividad",
    "API_Planilla", "API_Produccion",
    "data", "datos", "master",
    "Planilla Funcionario", "Produccion Funcionario",
]
for tbl in prod_tables:
    url = f"{PBI_BASE}/groups/{WS_ID}/datasets/{PROD_ID}/executeQueries"
    resp = requests.post(url, headers=H,
                         json={"queries": [{"query": f"EVALUATE TOPN(1, '{tbl}')"}],
                               "serializerSettings": {"includeNulls": True}},
                         timeout=10)
    if resp.ok:
        rows = resp.json().get("results",[{}])[0].get("tables",[{}])[0].get("rows",[])
        if rows:
            cols = list(rows[0].keys())
            print(f"  ✓ Tabla '{tbl}': {cols[:6]}")

print("\n=== FIN ===")
