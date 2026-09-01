#!/usr/bin/env python3
"""
Descubre TODOS los nombres de medidas en datasets de CxP, Productividad,
Planificacion y Consumo usando múltiples técnicas de la API de Power BI.
"""
import os, json, requests

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

def dax_raw(ds_id, query):
    url = f"{PBI_BASE}/groups/{WS_ID}/datasets/{ds_id}/executeQueries"
    resp = requests.post(url, headers=H,
                         json={"queries": [{"query": query}],
                               "serializerSettings": {"includeNulls": True}},
                         timeout=30)
    if resp.ok:
        rows = resp.json().get("results", [{}])[0].get("tables", [{}])[0].get("rows", [])
        return rows, None
    return [], resp.json().get("error", {})

DATASETS = {
    "cxp":            "45a8ab8d-e162-4398-a260-3a9a5f90829f",
    "productividad":  "a042ba6a-c82c-4fc1-bf95-b9e84bd15fc6",
    "planificacion":  "30074d92-7ec1-4762-82f2-1cb29c15dcfe",
    "consumo":        "c972c8cb-e5fc-4b60-8f5e-265a78e1e796",
}

# 1. Listar reportes del workspace para ver sus nombres
print("\n=== REPORTES EN WORKSPACE PAUNO ===")
reps = requests.get(f"{PBI_BASE}/groups/{WS_ID}/reports", headers=H).json().get("value", [])
for rep in reps:
    ds = rep.get("datasetId", "")
    ds_key = next((k for k, v in DATASETS.items() if v == ds), ds[:8])
    print(f"  [{ds_key}] {rep['name']} — id={rep['id']}")

# 2. Para cada dataset: intentar DMV queries para descubrir medidas
print("\n=== DMV QUERIES POR DATASET ===")
DMV_QUERIES = [
    ('INFO.MEASURES',
     'EVALUATE SELECTCOLUMNS(INFO.MEASURES(), "Tabla", [TableID], "Nombre", [Name], "Expr", [Expression])'),
    ('INFO.TABLES',
     'EVALUATE SELECTCOLUMNS(INFO.TABLES(), "ID", [ID], "Nombre", [Name])'),
    ('INFO.COLUMNS',
     'EVALUATE SELECTCOLUMNS(INFO.COLUMNS(), "Tabla", [TableID], "Nombre", [ExplicitName], "Tipo", [DataType])'),
    ('MDSCHEMA_MEASURES',
     'EVALUATE MDSCHEMA_MEASURES()'),
]

for ds_key, ds_id in DATASETS.items():
    print(f"\n--- {ds_key} ({ds_id[:8]}...) ---")
    for qname, q in DMV_QUERIES:
        rows, err = dax_raw(ds_id, q)
        if rows:
            print(f"  ✓ {qname}: {len(rows)} filas")
            for row in rows[:30]:
                print(f"    {row}")
        else:
            code = err.get("pbi.error", {}).get("code", "?") if err else "empty"
            print(f"  ✗ {qname}: {code}")

    # 3. Intentar EVALUATE sobre posibles nombres de tablas con schema completo
    print(f"  Schema REST API:")
    schema_r = requests.get(
        f"{PBI_BASE}/groups/{WS_ID}/datasets/{ds_id}/tables", headers=H)
    if schema_r.ok:
        for tbl in schema_r.json().get("value", []):
            cols = [c["name"] for c in tbl.get("columns", [])]
            meas = [m["name"] for m in tbl.get("measures", [])]
            print(f"    Tabla '{tbl['name']}': cols={cols[:5]}, medidas={meas[:10]}")
    else:
        print(f"    {schema_r.status_code}: {schema_r.text[:100]}")

    # 4. Probar EVALUATE ROW con nombres de medidas genericos
    generic = ["Total", "Saldo", "Importe", "Monto", "Valor", "Resultado",
               "KPI", "Indicador", "Dato", "Numero", "N", "V", "Medida"]
    print(f"  Medidas genéricas:")
    for m in generic:
        rows2, _ = dax_raw(ds_id, f'EVALUATE ROW("v", [{m}])')
        if rows2:
            v = list(rows2[0].values())[0] if rows2 else None
            print(f"    ✓ [{m}] = {v}")

print("\n=== FIN ===")
