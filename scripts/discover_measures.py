#!/usr/bin/env python3
"""
Descubre TODAS las medidas reales de cada dataset via schema REST + executeQueries.
Prueba cada medida con y sin filtro de mes anterior para ver cuáles responden.
"""
import os, requests, datetime

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

today = datetime.date.today()
PREV_MONTH = today.month - 1 if today.month > 1 else 12
PREV_YEAR  = today.year  if today.month > 1 else today.year - 1

DATASETS = {
    "margen":     "38076daa-d2cd-4a93-858a-82c0a4cf8cb6",
    "compras":    "06408938-8202-424e-80c0-b42c178dabde",
    "inventario": "0e27d784-41a4-48f0-9208-60210119f0a7",
    "mermas":     "35866214-f4da-45a3-a5a2-aa0c8caffe78",
    "consumo":    "c972c8cb-e5fc-4b60-8f5e-265a78e1e796",
    "planificacion": "30074d92-7ec1-4762-82f2-1cb29c15dcfe",
    "cxc":        "2eec70cd-0820-408f-b938-a2cd547b0c18",
    "cxp":        "45a8ab8d-e162-4398-a260-3a9a5f90829f",
}

DATE_CTX = {
    "margen":    ("Calendario", "Date"),
    "mermas":    ("Calendario", "Date"),
    "compras":   ("Calendario", "Fecha"),
    "inventario":("Calendario", "Date"),
    "consumo":   ("Calendario", "Date"),
    "planificacion": ("Calendario", "Date"),
}

def dax(ds_id, query):
    url = f"{PBI_BASE}/groups/{WS_ID}/datasets/{ds_id}/executeQueries"
    resp = requests.post(url, headers=H,
                         json={"queries": [{"query": query}],
                               "serializerSettings": {"includeNulls": True}},
                         timeout=20)
    if not resp.ok:
        return None, resp.text[:120]
    rows = resp.json().get("results",[{}])[0].get("tables",[{}])[0].get("rows",[])
    return rows, None

def get_schema(ds_id):
    url = f"{PBI_BASE}/groups/{WS_ID}/datasets/{ds_id}/tables"
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=20)
    if not r.ok:
        return {}
    schema = {}
    for tbl in r.json().get("value", []):
        name = tbl.get("name","")
        schema[name] = {
            "columns":  [c["name"] for c in tbl.get("columns",[])],
            "measures": [m["name"] for m in tbl.get("measures",[])],
        }
    return schema

def try_measure_plain(ds_id, m):
    rows, _ = dax(ds_id, f'EVALUATE ROW("v", [{m}])')
    if rows and rows[0]:
        return list(rows[0].values())[0]
    return "ERR"

def try_measure_dated(ds_id, m, tbl, col):
    q = f"""EVALUATE ROW("v", CALCULATE([{m}],
  FILTER(ALL('{tbl}'), YEAR('{tbl}'[{col}]) = {PREV_YEAR} && MONTH('{tbl}'[{col}]) = {PREV_MONTH})))"""
    rows, _ = dax(ds_id, q)
    if rows and rows[0]:
        return list(rows[0].values())[0]
    return "ERR"

print(f"=== DISCOVERY COMPLETO — {PREV_YEAR}-{PREV_MONTH:02d} (mes anterior) ===\n")

for ds_key, ds_id in DATASETS.items():
    print(f"\n{'='*60}")
    print(f"DATASET: {ds_key} ({ds_id[:8]}...)")
    date_ctx = DATE_CTX.get(ds_key)
    schema = get_schema(ds_id)
    if not schema:
        print("  ⚠ Schema vacío (Live Connection) — probando medidas por nombre...")
        CANDIDATOS_DS = {
            "margen":       ["Venta Total","Margen Variable","% Margen Variable","MV","% MV","MARGEN VARIABLE","Venta Neta","Precio/kg"],
            "compras":      ["Ratio","Cant Compra","Cant Consumo","Eficiencia","Eficiencia Costo","Ratio C/V","% Ratio","Consumo/Compra","Valor Compras","Total Compras","Monto Compras"],
            "inventario":   ["Stock Total","Dead Stock","% Dead Stock","Inventario Total","Working Stock","Rotacion","Dias Inventario"],
            "mermas":       ["% Merma Total","% Merma","Merma Total","Merma %","% Merma Ate","% Merma Pachacamac"],
            "consumo":      ["Consumo","Compras","Total Consumo","Ratio","% Ratio"],
            "planificacion":["% Avance","Avance","Ventas Real","Fill Rate","% Fill Rate"],
            "cxc":          ["% Morosidad","Morosidad","Por Vencer","CxC Total","Vencido"],
            "cxp":          ["Cuentas x Pagar","CUENTAS X PAGAR","Refinanciamiento","Proveedores"],
        }
        for m in CANDIDATOS_DS.get(ds_key, []):
            v = try_measure_plain(ds_id, m)
            if v != "ERR":
                if date_ctx:
                    v_d = try_measure_dated(ds_id, m, date_ctx[0], date_ctx[1])
                    print(f"  [{m}] sin filtro={v}  |  {PREV_YEAR}-{PREV_MONTH:02d}={v_d}")
                else:
                    print(f"  [{m}] = {v}")
        continue
    for tbl_name, tbl_info in schema.items():
        cols = tbl_info["columns"]
        meds = tbl_info["measures"]
        print(f"\n  Tabla: '{tbl_name}'")
        if cols: print(f"    Columnas: {cols[:8]}")
        if meds:
            print(f"    Medidas ({len(meds)}): {meds}")
            print(f"    --- Probando medidas ---")
            for m in meds:
                v_plain = try_measure_plain(ds_id, m)
                if v_plain == "ERR":
                    print(f"      [{m}] = no accesible")
                    continue
                if date_ctx:
                    v_dated = try_measure_dated(ds_id, m, date_ctx[0], date_ctx[1])
                    print(f"      [{m}] sin filtro={v_plain}  |  {PREV_YEAR}-{PREV_MONTH:02d}={v_dated}")
                else:
                    print(f"      [{m}] = {v_plain}")

print("\n=== FIN ===")
