#!/usr/bin/env python3
"""
Descubre TODAS las medidas reales de cada dataset via schema REST + executeQueries.
Guarda resultados en data/latest/discovered_measures.json para que fetch_powerbi.py
use los nombres exactos sin adivinar.
"""
import os, json, requests, datetime
from pathlib import Path

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
    "margen":        "38076daa-d2cd-4a93-858a-82c0a4cf8cb6",
    "compras":       "06408938-8202-424e-80c0-b42c178dabde",
    "inventario":    "0e27d784-41a4-48f0-9208-60210119f0a7",
    "mermas":        "35866214-f4da-45a3-a5a2-aa0c8caffe78",
    "consumo":       "c972c8cb-e5fc-4b60-8f5e-265a78e1e796",
    "planificacion": "30074d92-7ec1-4762-82f2-1cb29c15dcfe",
    "cxc":           "2eec70cd-0820-408f-b938-a2cd547b0c18",
    "cxp":           "45a8ab8d-e162-4398-a260-3a9a5f90829f",
    "fill_rate":     "7f4ebe22-5e90-4e35-973b-4af3c58497e5",
}

DATE_CTX = {
    "margen":        ("Calendario", "Date"),
    "mermas":        ("Calendario", "Date"),
    "compras":       ("Calendario", "Date"),
    "inventario":    ("Calendario", "Date"),
    "consumo":       ("Calendario", "Date"),
    "planificacion": ("Calendario", "Date"),
    "fill_rate":     ("Calendario", "Date"),
}

# Candidatos ampliados — todos los nombres posibles por dataset
CANDIDATOS = {
    "margen":   ["Venta Total","Ventas","Margen Variable","% Margen Variable","MV","% MV",
                 "MARGEN VARIABLE","Venta Neta","Precio/kg","Margen Bruto","% Margen"],
    "compras":  ["Ratio","Cant Compra","Cant Consumo","Eficiencia","Eficiencia Costo",
                 "Ratio C/V","% Ratio","Consumo/Compra","Valor Compras","Total Compras",
                 "Monto Compras","Compras","Importe Compras","Compra Mes","Consumo Mes",
                 "Monto Compra","Monto Consumo","Ratio Operativo","% Eficiencia"],
    "inventario":["Stock Total","Dead Stock","% Dead Stock","Inventario Total",
                  "Working Stock","Rotacion","Dias Inventario","Valor Stock",
                  "Stock Valorizado","Inmovilizado","% Inmovilizado"],
    "mermas":   ["% Merma Total","% Merma","Merma Total","Merma %",
                 "% Merma Ate","% Merma Pachacamac","Merma Kg","Merma Soles"],
    "consumo":  ["Consumo","Compras","Total Consumo","Total Compras","Ratio","% Ratio",
                 "Eficiencia","Cant Compra","Cant Consumo","Monto Consumo","Monto Compras",
                 "Consumo Periodo","Compra Periodo","Ratio C/C","Importe Consumo",
                 "Importe Compras","% Eficiencia","Ratio Operativo"],
    "planificacion":["% Avance","Avance","Ventas Real","Fill Rate","% Fill Rate",
                     "Presupuesto","PPTO","Avance vs Ppto","% Cumplimiento Ppto"],
    "cxc":      ["% Morosidad","Morosidad","Por Vencer","CxC Total","Vencido",
                 "Saldo CxC","Total CxC","% Vencido","Cartera Vencida"],
    "cxp":      ["Cuentas x Pagar","CUENTAS X PAGAR","Refinanciamiento","Proveedores",
                 "CxP Total","Total CxP","Saldo CxP","Deuda Total"],
    "fill_rate":["% FILLRATE","% Fill Rate","ORDEN DE VENTA","FACTURACION","VENTA PERDIDA",
                 "Fill Rate","Orden de Venta","Facturacion","Venta Perdida",
                 "% Cumplimiento","Pedidos","Atendidos","No Atendidos"],
}

def dax(ds_id, query):
    import time
    url = f"{PBI_BASE}/groups/{WS_ID}/datasets/{ds_id}/executeQueries"
    for attempt in range(2):
        resp = requests.post(url, headers=H,
                             json={"queries": [{"query": query}],
                                   "serializerSettings": {"includeNulls": True}},
                             timeout=25)
        if resp.status_code == 429:
            time.sleep(int(resp.headers.get("Retry-After", 15)))
            continue
        if not resp.ok:
            return None
        rows = resp.json().get("results",[{}])[0].get("tables",[{}])[0].get("rows",[])
        return rows
    return None

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

def try_plain(ds_id, m):
    rows = dax(ds_id, f'EVALUATE ROW("v", [{m}])')
    if rows and rows[0]:
        v = list(rows[0].values())[0]
        return v
    return None

def try_dated(ds_id, m, tbl, col):
    q = f"""EVALUATE ROW("v", CALCULATE([{m}],
  FILTER(ALL('{tbl}'), YEAR('{tbl}'[{col}]) = {PREV_YEAR} && MONTH('{tbl}'[{col}]) = {PREV_MONTH})))"""
    rows = dax(ds_id, q)
    if rows and rows[0]:
        return list(rows[0].values())[0]
    return None

print(f"=== DISCOVERY COMPLETO — {PREV_YEAR}-{PREV_MONTH:02d} ===\n")

results = {}  # {ds_key: {measure_name: {plain: val, dated: val}}}

for ds_key, ds_id in DATASETS.items():
    print(f"\n{'='*60}")
    print(f"DATASET: {ds_key} ({ds_id[:8]}...)")
    date_ctx = DATE_CTX.get(ds_key)
    results[ds_key] = {}

    # Intentar schema REST primero (funciona solo en datasets no-Live-Connection)
    schema = get_schema(ds_id)
    all_measures = []
    if schema:
        for tbl_name, tbl_info in schema.items():
            meds = tbl_info.get("measures", [])
            if meds:
                print(f"  Tabla '{tbl_name}': {len(meds)} medidas → {meds}")
                all_measures.extend(meds)
    else:
        print("  Schema vacío (Live Connection) — usando candidatos conocidos")
        all_measures = CANDIDATOS.get(ds_key, [])

    print(f"  Probando {len(all_measures)} medidas...")
    for m in all_measures:
        v_plain = try_plain(ds_id, m)
        if v_plain is None:
            continue
        entry = {"plain": v_plain}
        if date_ctx:
            v_dated = try_dated(ds_id, m, date_ctx[0], date_ctx[1])
            entry["dated"] = v_dated
            print(f"  ✓ [{m}]  sin_filtro={v_plain}  |  agosto={v_dated}")
        else:
            print(f"  ✓ [{m}] = {v_plain}")
        results[ds_key][m] = entry

# Guardar en JSON para que fetch_powerbi.py lo use directamente
out_path = Path("data/latest/discovered_measures.json")
out_path.parent.mkdir(parents=True, exist_ok=True)
out_path.write_text(json.dumps({
    "generado": str(today),
    "mes_filtro": f"{PREV_YEAR}-{PREV_MONTH:02d}",
    "datasets": results
}, ensure_ascii=False, indent=2))
print(f"\n✓ Guardado en {out_path}")
print("\n=== FIN ===")
