#!/usr/bin/env python3
"""
JUANITO — Power BI Data Fetcher v3
Estrategia: scan de medidas reales → query con nombres confirmados.
"""

import os, json, requests, datetime, sys
from pathlib import Path

try:
    import warnings; warnings.filterwarnings('ignore')
except: pass

TENANT_ID     = os.environ["AZURE_TENANT_ID"]
CLIENT_ID     = os.environ["AZURE_CLIENT_ID"]
CLIENT_SECRET = os.environ["AZURE_CLIENT_SECRET"]
USERNAME      = os.environ["PBI_USERNAME"]
PASSWORD      = os.environ["PBI_PASSWORD"]

PBI_SCOPE = "https://analysis.windows.net/powerbi/api/.default"
TOKEN_URL = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
PBI_BASE  = "https://api.powerbi.com/v1.0/myorg"

OUTPUT_DIR = Path("data/latest")
HIST_DIR   = Path("data/historico")
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
HIST_DIR.mkdir(parents=True, exist_ok=True)

HOY = datetime.date.today().strftime("%Y-%m-%d")

WORKSPACES = {
    "PAUNO": "461932ad-b5ec-4fd6-aa97-f1fc7bdc5169",
}

# IDs confirmados por discover_all_datasets() en ejecución anterior
DATASET_IDS = {
    "PAUNO": {
        "cxc":            "2eec70cd-0820-408f-b938-a2cd547b0c18",  # 1. Cuentas por cobrar
        "cxp":            "45a8ab8d-e162-4398-a260-3a9a5f90829f",  # 2. Cuentas por pagar
        "margen":         "38076daa-d2cd-4a93-858a-82c0a4cf8cb6",  # 3. Reporte de Margen
        "mermas":         "35866214-f4da-45a3-a5a2-aa0c8caffe78",  # 4. Reporte de mermas
        "compras":        "06408938-8202-424e-80c0-b42c178dabde",  # 5. Reporte de compras
        "inventario":     "0e27d784-41a4-48f0-9208-60210119f0a7",  # 6. Rotacion de inventario
        "control_ds":     "0aca7bdd-6b72-49c2-be41-17ae0f6b5848",  # 8. Reporte de auditoria
        "productividad_ds": "a042ba6a-c82c-4fc1-bf95-b9e84bd15fc6", # 13. Productividad
        "planificacion":  "30074d92-7ec1-4762-82f2-1cb29c15dcfe",  # 11. Planificaciones
        "consumo":        "c972c8cb-e5fc-4b60-8f5e-265a78e1e796",  # 14. Consumo Materiales
    }
}

# Candidatos de medidas por tipo — orden de más probable a menos
SCAN_CANDIDATES = {
    "cxc": [
        "% Morosidad", "Morosidad", "% Mora", "Mora", "Tasa Morosidad",
        "Por Vencer", "Saldo Por Vencer", "CxC Por Vencer", "No Vencido",
        "CxC Total", "Total CxC", "Saldo CxC", "Cartera Total",
        "Vencido", "Saldo Vencido", "Total Vencido",
        "0-15 días", "16-30 días", "+30 días", "0-15d", "16-30d", "+30d",
        "Rotacion CxC", "Días CxC", "Días de Cobro",
    ],
    "cxp": [
        "DPP", "Días de Pago", "Dias de Pago", "Días CxP", "Dias CxP",
        "Plazo Pago", "Plazo Promedio Pago", "Días Promedio de Pago",
        "Días Promedio de Pago Actual", "Días Promedio",
        "CxP Total", "Total CxP", "Saldo CxP", "Saldo Proveedores",
        "Deuda Total", "Importe CxP", "Total Proveedores",
        "Vencido CxP", "Por Vencer CxP", "Saldo Pagar",
        "Refinanciado", "Saldo Refinanciado",
        "Monto CxP", "Total Deuda", "CxP",
        "Días de Crédito", "Rotacion CxP", "Plazo Crédito",
        # Variantes sin tilde/acento
        "Dias Promedio Pago", "Dias de Credito", "Dias Credito",
        "Total Pagar", "Saldo Total", "Periodo Promedio Pago",
        "PPP", "PPagos", "Plazo Promedio",
        "% Vencido", "Vencido Total", "Vencimiento",
    ],
    "margen": [
        "Margen Variable", "% Margen Variable", "MV", "Margen Bruto",
        "% MV", "MV%", "Margen", "% Margen",
        "Ventas Mes Actual", "Ventas Actuales", "Ventas", "Venta Total", "Venta Neta",
        "Ventas Mes Anterior", "Ventas Anterior",
        "Precio/kg", "Precio kg", "Precio Promedio kg",
        "Costo Variable", "% Costo Variable",
    ],
    "mermas": [
        "% Merma", "Merma %", "Tasa Merma", "Merma Total", "% Merma Total",
        "% Merma Ate", "Merma Ate", "Merma Planta Ate",
        "% Merma Pachacamac", "Merma Pachacamac", "Merma Lurin",
        "% Merma Terceros", "Merma Terceros",
        "Merma Valorizada", "Kg Merma",
    ],
    "compras": [
        "Ratio", "Ratio C/V", "Ratio Compras", "Ratio Consumo/Compra",
        "% Ratio", "Ratio C/C", "Consumo/Compra", "C/C",
        "Consumo Total", "Consumo Periodo", "Importe Consumo", "Monto Consumo",
        "Compra Total", "Monto Compras", "Importe Compras", "Valor Compras",
        "MP Consumo", "MP Compras", "Materia Prima Consumo",
        "Total Compras Periodo", "Reposicion", "Ordenes de Compra",
        # Variantes adicionales
        "Consumo vs Compras", "Eficiencia Compras", "% Eficiencia",
        "Compra Mes", "Consumo Mes", "Compras Mes",
        "Valor Compra", "Valor Consumo",
        "Total", "Importe Total",
        "Compra Neta", "Compra Bruta",
        "Req Compras", "Requerimiento",
    ],
    "inventario": [
        "Dead Stock", "Stock Muerto", "Inmovilizado", "Stock Inmovilizado",
        "% Dead Stock", "% Inmovilizado", "% Dead", "Pct Dead",
        "Inventario Total", "Total Inventario", "Saldo Inventario", "Valor Inventario",
        "Working Stock", "Stock Activo", "Stock Working", "Stock Normal",
        "Días de Inventario", "Dias Inventario", "Cobertura", "Cobertura Días",
        "Exceso 1", "Exceso 2", "Sobre Stock",
        "Rotacion", "Rotación", "Veces Rotacion",
        "Stock Total", "Total Stock", "Costo Inventario",
    ],
    "control_ds": [
        "% Cumplimiento", "Cumplimiento", "Conformidad", "% Conformidad",
        "Puntos Críticos", "Puntos Criticos", "Hallazgos Críticos",
        "Satisfactorio", "Observaciones",
        "No Conformidades", "Incumplimientos",
    ],
    "productividad_ds": [
        "Planilla/kg", "S//kg", "Costo/kg", "Costo Planilla kg",
        "Productividad", "Planilla Total", "Costo Planilla",
        "Kg Producidos", "Kg Producción", "Produccion Total",
        "Eficiencia", "HH/kg", "Costo por kg", "S/ por kg",
        "Planilla Mensual", "Total Planilla", "Gasto Planilla",
        "Kg Produccion", "Produccion Kg", "Total Kg",
        "Ratio Planilla", "Costo Mano de Obra",
        "MOD", "MOD/kg", "Mano de Obra",
    ],
    "planificacion": [
        "Fill Rate", "% Fill Rate", "Tasa Atención", "Tasa Atencion",
        "% Atención", "% Atencion",
        "Dead Stock", "% Dead Stock",
        "Avance", "% Avance", "Avance PPTO", "% Avance Presupuesto",
        "Presupuesto", "PPTO", "Ventas PPTO",
    ],
    "consumo": [
        "Consumo", "Total Consumo", "Importe Consumo",
        "Compras", "Total Compras",
        "Ratio", "Ratio Consumo", "% Ratio",
    ],
}

# ─────────────────────────────────────────────────────────────────────────────

def get_token():
    r = requests.post(TOKEN_URL, data={
        "grant_type": "password", "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET, "username": USERNAME,
        "password": PASSWORD, "scope": PBI_SCOPE,
    }, timeout=30)
    r.raise_for_status()
    return r.json()["access_token"]

def try_measure(token, ws_id, dataset_id, name):
    """Prueba si una medida existe. Retorna (valor, True) o (None, False)."""
    r = requests.post(
        f"{PBI_BASE}/groups/{ws_id}/datasets/{dataset_id}/executeQueries",
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json={"queries": [{"query": f'EVALUATE ROW("v", [{name}])'}],
              "serializerSettings": {"includeNulls": True}},
        timeout=20,
    )
    if not r.ok:
        return None, False
    rows = r.json().get("results", [{}])[0].get("tables", [{}])[0].get("rows", [])
    v = (rows[0].get("[v]") or rows[0].get("v")) if rows else None
    return v, True

MEASURE_CACHE_FILE = OUTPUT_DIR / "measure_cache.json"

def load_measure_cache():
    try:
        if MEASURE_CACHE_FILE.exists():
            return json.loads(MEASURE_CACHE_FILE.read_text())
    except: pass
    return {}

def save_measure_cache(cache):
    try:
        MEASURE_CACHE_FILE.write_text(json.dumps(cache, ensure_ascii=False, indent=2))
    except: pass

def scan_dataset(token, ws_id, dataset_id, ds_key, cache=None):
    """Escanea medidas reales. Usa caché para omitir candidatos ya descartados."""
    candidates = SCAN_CANDIDATES.get(ds_key, [])
    found = {}
    # Medidas ya confirmadas en caché
    cached_found = (cache or {}).get(ds_key, {})
    if cached_found:
        # Verificar que las medidas cacheadas siguen funcionando
        print(f"    Verificando {len(cached_found)} medidas en caché...")
        for name in list(cached_found.keys()):
            v, exists = try_measure(token, ws_id, dataset_id, name)
            if exists:
                found[name] = v
                print(f"      ✓ [{name}] = {v} (caché)")
        # Solo scan nuevos candidatos no cacheados
        scanned_before = set(cached_found.keys())
        candidates = [c for c in candidates if c not in scanned_before]
    print(f"    Escaneando {len(candidates)} candidatos nuevos...")
    for name in candidates:
        v, exists = try_measure(token, ws_id, dataset_id, name)
        if exists:
            found[name] = v
            print(f"      ✓ [{name}] = {v}")
    return found

def discover_tables_in_dataset(token, ws_id, dataset_id):
    """Descubre tablas del dataset probando nombres comunes. Retorna {table_name: [cols]}."""
    common_tables = [
        "Proveedores", "Facturas", "CxP", "Comprobantes", "Pagos",
        "Cuentas por Pagar", "Tabla CxP", "Detalle CxP",
        "Compras", "Ordenes de Compra", "OC", "Materiales",
        "Consumo", "Tabla Compras", "Detalle Compras",
        "Mermas", "Produccion", "Planilla", "Empleados",
        "Ventas", "Clientes", "Facturas Ventas",
        "Inventario", "Stock", "Articulos", "Productos",
        "Control Interno", "Auditoría", "Auditoria",
        "Medidas", "_Medidas", "KPIs",
        "Calendario", "Fecha", "Calendar",
        "fCalendario", "dCalendario",
    ]
    found_tables = {}
    for tbl in common_tables:
        r = requests.post(
            f"{PBI_BASE}/groups/{ws_id}/datasets/{dataset_id}/executeQueries",
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json={"queries": [{"query": f"EVALUATE TOPN(3, '{tbl}')"}],
                  "serializerSettings": {"includeNulls": True}},
            timeout=15,
        )
        if r.ok:
            rows = r.json().get("results", [{}])[0].get("tables", [{}])[0].get("rows", [])
            if rows:
                cols = list(rows[0].keys())
                found_tables[tbl] = cols
                print(f"    Tabla '{tbl}': {cols[:5]}")
    return found_tables

def discover_all_datasets(token, ws_id):
    r = requests.get(f"{PBI_BASE}/groups/{ws_id}/datasets",
                     headers={"Authorization": f"Bearer {token}"}, timeout=20)
    if not r.ok: return {}
    return {d["name"]: d["id"] for d in r.json().get("value", [])}

def fmt_soles(v, decimals=0):
    if v is None: return "—"
    try:
        n = float(str(v).replace(",", ".").replace("%", "").replace("S/", "").replace(" ", ""))
        if abs(n) >= 1_000_000: return f"S/{n/1_000_000:.2f}M"
        elif abs(n) >= 1_000: return f"S/{n:,.0f}"
        return f"S/{n:.{decimals}f}"
    except: return str(v)

def fmt_pct(v):
    if v is None: return "—"
    try:
        n = float(str(v).replace(",", ".").replace("%", ""))
        if abs(n) < 1: n *= 100
        return f"{n:.1f}%"
    except: return str(v)

def to_float(v):
    if v is None: return None
    try:
        n = float(str(v).replace(",", ".").replace("%", ""))
        return n
    except: return None

def sem_thresh(v, red_above=None, yellow_above=None, red_below=None, yellow_below=None):
    n = to_float(v)
    if n is None: return "green"
    if abs(n) < 1 and (red_above or yellow_above or red_below or yellow_below):
        n *= 100
    if red_above and n >= red_above: return "red"
    if yellow_above and n >= yellow_above: return "yellow"
    if red_below and n <= red_below: return "red"
    if yellow_below and n <= yellow_below: return "yellow"
    return "green"

# ─── Funciones de reporte ─────────────────────────────────────────────────────

def build_cxc(found):
    mora_val = found.get("% Morosidad") or found.get("Morosidad") or found.get("% Mora") or found.get("Mora")
    vencer_val = found.get("Por Vencer") or found.get("Saldo Por Vencer") or found.get("No Vencido")
    total_val = found.get("CxC Total") or found.get("Total CxC") or found.get("Saldo CxC") or found.get("Cartera Total")
    vencido_val = found.get("Vencido") or found.get("Saldo Vencido") or found.get("Total Vencido")

    mora_pct = to_float(mora_val)
    if mora_pct and abs(mora_pct) < 1: mora_pct *= 100
    sem = sem_thresh(mora_pct, red_above=20, yellow_above=15)

    kpis = []
    if mora_pct is not None:
        kpis.append({"label": "Morosidad", "valor": f"{mora_pct:.1f}%", "meta": "15%", "estado": sem})
    if vencer_val is not None:
        kpis.append({"label": "CxC por vencer", "valor": fmt_soles(vencer_val)})
    if total_val is not None:
        kpis.append({"label": "CxC Total", "valor": fmt_soles(total_val)})
    if vencido_val is not None:
        kpis.append({"label": "CxC Vencido", "valor": fmt_soles(vencido_val), "estado": "red" if sem == "red" else "yellow"})

    razon = f"Morosidad {fmt_pct(mora_pct)}" + (" — crítica" if sem=="red" else " — sobre meta" if sem=="yellow" else " — OK")
    alerta = razon if sem != "green" else None

    tramos = []
    for label, key in [("Por vencer", "Por Vencer"), ("0-15d", "0-15 días"), ("16-30d", "16-30 días"), ("+30d", "+30 días")]:
        v = found.get(key) or found.get(key.replace(" días", "d"))
        if v is not None:
            s = "green" if "vencer" in label.lower() else ("yellow" if "15" in label else "red")
            tramos.append({"label": label, "valor": fmt_soles(v), "estado": s})

    result = {"estado": sem, "alerta": alerta, "kpis": kpis}
    if tramos: result["tramos"] = tramos
    return result, sem, razon, mora_pct, vencer_val

def build_cxp(found):
    dias_val = (found.get("DPP") or found.get("Días de Pago") or found.get("Dias de Pago") or
                found.get("Días CxP") or found.get("Dias CxP") or found.get("Plazo Pago") or
                found.get("Plazo Promedio Pago") or found.get("Días Promedio de Pago"))
    total_val = (found.get("CxP Total") or found.get("Total CxP") or found.get("Saldo CxP") or
                 found.get("Saldo Proveedores") or found.get("Deuda Total"))
    refin_val = found.get("Refinanciado") or found.get("Saldo Refinanciado")

    dias = to_float(dias_val)
    sem = "green"
    if dias:
        if dias > 90: sem = "red"
        elif dias > 60: sem = "yellow"

    kpis = []
    if dias is not None:
        kpis.append({"label": "Días CxP", "valor": f"{dias:.0f} días", "meta": "<90d", "estado": sem})
    if total_val is not None:
        kpis.append({"label": "CxP Total", "valor": fmt_soles(total_val)})
    if refin_val is not None:
        kpis.append({"label": "Refinanciado", "valor": fmt_soles(refin_val)})

    alerta = f"CxP {dias:.0f}d — revisar flujo" if sem != "green" and dias else None
    return {"estado": sem, "alerta": alerta, "kpis": kpis}, dias

def build_margen(found):
    margen_val = (found.get("Margen Variable") or found.get("% Margen Variable") or
                  found.get("MV") or found.get("% MV") or found.get("Margen") or found.get("% Margen"))
    ventas_val = (found.get("Ventas Mes Actual") or found.get("Ventas Actuales") or
                  found.get("Ventas") or found.get("Venta Total") or found.get("Venta Neta"))
    precio_val = found.get("Precio/kg") or found.get("Precio kg") or found.get("Precio Promedio kg")

    margen_pct = to_float(margen_val)
    if margen_pct and abs(margen_pct) < 1: margen_pct *= 100
    sem = sem_thresh(margen_pct, red_below=46, yellow_below=52)

    kpis = []
    if ventas_val is not None:
        kpis.append({"label": "Ventas mes", "valor": fmt_soles(ventas_val)})
    if margen_pct is not None:
        kpis.append({"label": "Margen variable", "valor": f"{margen_pct:.1f}%", "meta": "52%", "estado": sem})
    if precio_val is not None:
        kpis.append({"label": "Precio/kg", "valor": f"S/{to_float(precio_val):.2f}"})

    alerta = f"Margen {fmt_pct(margen_pct)} — bajo meta 52%" if sem != "green" and margen_pct else None
    return {"estado": sem, "alerta": alerta, "kpis": kpis}, ventas_val, margen_pct

def build_mermas(found):
    ate_val  = found.get("% Merma Ate") or found.get("Merma Ate") or found.get("Merma Planta Ate")
    pach_val = found.get("% Merma Pachacamac") or found.get("Merma Pachacamac") or found.get("Merma Lurin")
    tot_val  = found.get("% Merma") or found.get("Merma %") or found.get("Tasa Merma") or found.get("Merma Total") or found.get("% Merma Total")
    terc_val = found.get("% Merma Terceros") or found.get("Merma Terceros")

    def pct(v):
        n = to_float(v)
        if n is None: return None
        if abs(n) < 1: n *= 100
        return n

    ate, pach, tot, terc = pct(ate_val), pct(pach_val), pct(tot_val), pct(terc_val)
    worst = max(filter(lambda x: x is not None, [ate, pach, tot, terc]), default=None)
    sem = "green"
    if worst:
        if worst > 3: sem = "red"
        elif worst > 2: sem = "yellow"

    kpis = []
    for label, val in [("Merma Ate", ate), ("Merma Pachacamac", pach), ("Merma Terceros", terc), ("Merma Total", tot)]:
        if val is not None:
            s = "red" if val > 3 else ("yellow" if val > 2 else "green")
            kpis.append({"label": label, "valor": f"{val:.2f}%", "meta": "2%", "estado": s})

    alerta = f"Merma {worst:.2f}% — sobre meta 2%" if sem != "green" and worst else None
    return {"estado": sem, "alerta": alerta, "kpis": kpis}

def build_compras(found):
    ratio_val  = (found.get("Ratio") or found.get("Ratio C/V") or found.get("Ratio Compras") or
                  found.get("Ratio Consumo/Compra") or found.get("% Ratio") or found.get("Consumo/Compra"))
    consumo_val = found.get("Consumo") or found.get("Total Consumo") or found.get("Importe Consumo")
    compra_val  = found.get("Compras") or found.get("Total Compras") or found.get("Importe Compras") or found.get("Valor Compras")

    ratio = to_float(ratio_val)
    if ratio and abs(ratio) < 2: ratio *= 100
    if ratio is None and consumo_val and compra_val:
        c, p = to_float(consumo_val), to_float(compra_val)
        if c and p and p > 0: ratio = (c / p) * 100

    sem = "green"
    if ratio:
        if ratio > 130 or ratio < 70: sem = "red"
        elif ratio > 110 or ratio < 80: sem = "yellow"

    interp = ("jala inventario" if ratio and ratio > 100 else
              "sobre-compra" if ratio and ratio < 80 else "normal")
    kpis = []
    if ratio is not None:
        kpis.append({"label": "Ratio Consumo/Compra", "valor": f"{ratio:.1f}%", "meta": "80-100%",
                     "estado": sem, "interpretacion": interp})
    if consumo_val is not None:
        kpis.append({"label": "Consumo", "valor": fmt_soles(consumo_val)})
    if compra_val is not None:
        kpis.append({"label": "Compras", "valor": fmt_soles(compra_val)})

    alerta = f"Ratio {ratio:.1f}% — {interp}" if sem != "green" and ratio else None
    return {"estado": sem, "alerta": alerta, "kpis": kpis}, ratio

def build_inventario(found):
    dead_val    = found.get("Dead Stock") or found.get("Stock Muerto") or found.get("Inmovilizado") or found.get("Stock Inmovilizado")
    deadpct_val = found.get("% Dead Stock") or found.get("% Inmovilizado") or found.get("% Dead") or found.get("Pct Dead")
    total_val   = (found.get("Inventario Total") or found.get("Total Inventario") or
                   found.get("Saldo Inventario") or found.get("Valor Inventario") or
                   found.get("Stock Total") or found.get("Total Stock") or found.get("Costo Inventario"))
    working_val = found.get("Working Stock") or found.get("Stock Activo") or found.get("Stock Working") or found.get("Stock Normal")

    dead_pct = to_float(deadpct_val)
    if dead_pct and abs(dead_pct) < 1: dead_pct *= 100
    if dead_pct is None and dead_val and total_val:
        d, t = to_float(dead_val), to_float(total_val)
        if d and t and t > 0: dead_pct = d / t * 100

    sem = sem_thresh(dead_pct, red_above=10, yellow_above=5)

    kpis = []
    if dead_val is not None:
        kpis.append({"label": "Dead Stock", "valor": fmt_soles(dead_val), "meta": "<S/500K", "estado": sem})
    if dead_pct is not None:
        kpis.append({"label": "% Dead Stock", "valor": f"{dead_pct:.1f}%", "meta": "<5%", "estado": sem})
    if total_val is not None:
        kpis.append({"label": "Inventario Total", "valor": fmt_soles(total_val)})
    if working_val is not None:
        kpis.append({"label": "Working Stock", "valor": fmt_soles(working_val)})

    alerta = f"Dead Stock {dead_pct:.1f}% del inventario" if sem != "green" and dead_pct else None
    return {"estado": sem, "alerta": alerta, "kpis": kpis}

def build_control(found):
    cumpl_val  = found.get("% Cumplimiento") or found.get("Cumplimiento") or found.get("Conformidad") or found.get("% Conformidad")
    crit_val   = (found.get("Puntos Críticos") or found.get("Puntos Criticos") or
                  found.get("Hallazgos Críticos") or found.get("No Conformidades") or found.get("Incumplimientos"))
    satisf_val = found.get("Satisfactorio") or found.get("Observaciones")

    cumpl_pct = to_float(cumpl_val)
    if cumpl_pct and abs(cumpl_pct) < 1: cumpl_pct *= 100
    sem = sem_thresh(cumpl_pct, red_below=70, yellow_below=85)

    kpis = []
    if cumpl_pct is not None:
        kpis.append({"label": "Cumplimiento", "valor": f"{cumpl_pct:.1f}%", "meta": "85%", "estado": sem})
    if crit_val is not None:
        s = "red" if to_float(crit_val) and to_float(crit_val) > 0 else "green"
        kpis.append({"label": "Puntos críticos", "valor": str(int(to_float(crit_val) or 0)), "meta": "0", "estado": s})
    if satisf_val is not None:
        kpis.append({"label": "Satisfactorio", "valor": str(int(to_float(satisf_val) or satisf_val)), "estado": "green"})

    alerta = f"Cumplimiento {fmt_pct(cumpl_pct)} — bajo meta 85%" if sem != "green" and cumpl_pct else None
    return {"estado": sem if kpis else "green", "alerta": alerta, "kpis": kpis}

def build_productividad(found):
    pkg_val  = found.get("Planilla/kg") or found.get("S//kg") or found.get("Costo/kg") or found.get("Productividad")
    plan_val = found.get("Planilla Total") or found.get("Costo Planilla")
    kg_val   = found.get("Kg Producidos") or found.get("Kg Producción") or found.get("Produccion Total")

    kpis = []
    if pkg_val is not None:
        n = to_float(pkg_val)
        kpis.append({"label": "S/kg producido", "valor": f"S/{n:.2f}" if n else "—"})
    if plan_val is not None:
        kpis.append({"label": "Planilla total", "valor": fmt_soles(plan_val)})
    if kg_val is not None:
        n = to_float(kg_val)
        kpis.append({"label": "Kg producidos", "valor": f"{n:,.0f} kg" if n else "—"})

    return {"estado": "green", "alerta": None, "kpis": kpis}

def build_fill_rate(found):
    fill_val = found.get("Fill Rate") or found.get("% Fill Rate") or found.get("Tasa Atención") or found.get("Tasa Atencion")
    fill_pct = to_float(fill_val)
    if fill_pct and abs(fill_pct) < 1: fill_pct *= 100
    sem = sem_thresh(fill_pct, red_below=85, yellow_below=95)

    kpis = []
    if fill_pct is not None:
        kpis.append({"label": "Fill Rate", "valor": f"{fill_pct:.1f}%", "meta": "98%", "estado": sem})
        if fill_pct < 98:
            kpis.append({"label": "Brecha", "valor": f"-{98-fill_pct:.1f}pp"})

    alerta = f"Fill Rate {fmt_pct(fill_pct)} — bajo meta 98%" if sem != "green" and fill_pct else None
    return {"estado": sem, "alerta": alerta, "kpis": kpis}

def build_avance(found):
    avance_val = (found.get("Avance") or found.get("% Avance") or found.get("Avance PPTO") or
                  found.get("% Avance Presupuesto"))
    real_val   = found.get("Ventas PPTO") or found.get("Presupuesto")
    ppto_val   = found.get("Presupuesto") or found.get("PPTO")

    avance_pct = to_float(avance_val)
    if avance_pct and abs(avance_pct) < 1: avance_pct *= 100
    sem = sem_thresh(avance_pct, red_below=60, yellow_below=80)

    kpis = []
    if avance_pct is not None:
        kpis.append({"label": "Avance vs Ppto", "valor": f"{avance_pct:.1f}%", "meta": "100%", "estado": sem})
    if real_val is not None:
        kpis.append({"label": "Ventas reales", "valor": fmt_soles(real_val)})

    alerta = f"Avance {fmt_pct(avance_pct)} vs presupuesto" if sem != "green" and avance_pct else None
    return {"estado": sem, "alerta": alerta, "kpis": kpis}, avance_pct

# ─── Main ────────────────────────────────────────────────────────────────────

def main():
    print("=== JUANITO Power BI Fetcher v3 ===")
    print(f"Fecha: {HOY}")

    print("\n[1] Autenticando...")
    try:
        token = get_token()
        print("    Token OK")
    except Exception as e:
        print(f"    ERROR: {e}"); sys.exit(1)

    summary = {
        "fecha": datetime.date.today().strftime("%d/%m/%Y"),
        "fecha_actualizacion": HOY,
        "generado_por": "JUANITO — Power BI Direct v3",
        "semaforos": {}, "semaforo_razon": {},
        "holding_ventas": "—", "holding_mora": "—",
        "agenda_ceo": {"decidir_hoy": [], "escalar_semana": [], "monitorear": []},
        "empresas": {}
    }

    holding_ventas = 0
    holding_mora_vals = []

    for empresa, ws_id in WORKSPACES.items():
        print(f"\n[{empresa}]")
        ids = DATASET_IDS.get(empresa, {})
        empresa_data = {"reportes": {}}

        # Escanear medidas reales de cada dataset
        measure_cache = load_measure_cache()
        scanned = {}
        for ds_key, did in ids.items():
            print(f"  Escaneando {ds_key}...")
            scanned[ds_key] = scan_dataset(token, ws_id, did, ds_key, measure_cache)

        # Para datasets sin medidas: descubrir tablas y columnas
        for ds_key in ["cxp", "compras", "productividad_ds"]:
            if ds_key in ids and not scanned.get(ds_key):
                print(f"  Descubriendo tablas en {ds_key}...")
                tables = discover_tables_in_dataset(token, ws_id, ids[ds_key])
                if tables:
                    print(f"  {ds_key}: tablas encontradas = {list(tables.keys())}")
                    # Intentar usarlas para armar KPIs simples
                    scanned[ds_key]["_tables"] = tables

        # Guardar caché de medidas confirmadas
        new_cache = {}
        for ds_key, found in scanned.items():
            if found:
                new_cache[ds_key] = {k: None for k in found.keys()}  # guardar nombres, no valores
        save_measure_cache(new_cache)

        # Combinamos consumo en compras si aplica
        if "consumo" in scanned and "compras" in scanned:
            scanned["compras"].update(scanned.get("consumo", {}))

        # Combinamos planificacion con inventario
        if "planificacion" in scanned and "inventario" in scanned:
            scanned["inventario"].update(scanned.get("planificacion", {}))

        # ── CxC
        if scanned.get("cxc"):
            r, sem, razon, mora_pct, vencer = build_cxc(scanned["cxc"])
            empresa_data["reportes"]["cuentas_por_cobrar"] = r
            summary["semaforos"][empresa] = sem
            summary["semaforo_razon"][empresa] = razon
            if mora_pct: holding_mora_vals.append(mora_pct)
            mora_str = f"{mora_pct:.1f}" if mora_pct is not None else "—"
            print(f"  CxC mora={mora_str}% sem={sem}")
            if sem == "red":
                summary["agenda_ceo"]["decidir_hoy"].append({
                    "empresa": empresa,
                    "texto": f"Mora {fmt_pct(mora_pct)} — vencido {fmt_soles(vencer)}. Activar cobranza.",
                    "responsable": "Gerencia Financiera"
                })
        else:
            summary["semaforos"][empresa] = "yellow"
            summary["semaforo_razon"][empresa] = "Sin datos CxC"

        # ── CxP
        if scanned.get("cxp"):
            r, dias = build_cxp(scanned["cxp"])
            empresa_data["reportes"]["cuentas_por_pagar"] = r
            print(f"  CxP dias={dias} sem={r['estado']}")
            if r["estado"] == "red":
                summary["agenda_ceo"]["escalar_semana"].append({
                    "empresa": empresa, "texto": f"CxP {dias:.0f}d — riesgo con proveedores.", "responsable": "Finanzas"
                })

        # ── Margen
        if scanned.get("margen"):
            r, ventas, margen = build_margen(scanned["margen"])
            empresa_data["reportes"]["margen_variable"] = r
            if ventas:
                try: holding_ventas += float(ventas)
                except: pass
            print(f"  Margen {fmt_pct(margen)} ventas={fmt_soles(ventas)} sem={r['estado']}")
            if r["estado"] == "red":
                summary["agenda_ceo"]["escalar_semana"].append({
                    "empresa": empresa, "texto": f"Margen {fmt_pct(margen)} bajo meta 52%.", "responsable": "Comercial"
                })

        # ── Mermas
        if scanned.get("mermas"):
            r = build_mermas(scanned["mermas"])
            empresa_data["reportes"]["mermas"] = r
            print(f"  Mermas {len(r['kpis'])} KPIs sem={r['estado']}")

        # ── Compras
        if scanned.get("compras"):
            r, ratio = build_compras(scanned["compras"])
            empresa_data["reportes"]["compras"] = r
            print(f"  Compras ratio={ratio} sem={r['estado']}")
            if r["estado"] == "red":
                summary["agenda_ceo"]["decidir_hoy"].append({
                    "empresa": empresa,
                    "texto": f"Ratio {ratio:.1f}% — {'jalando inventario' if ratio and ratio>100 else 'sobre-compra'}.",
                    "responsable": "Logística"
                })

        # ── Inventario
        if scanned.get("inventario"):
            r = build_inventario(scanned["inventario"])
            empresa_data["reportes"]["sop_inventario"] = r
            print(f"  S&OP {len(r['kpis'])} KPIs sem={r['estado']}")

        # ── Control
        if scanned.get("control_ds"):
            r = build_control(scanned["control_ds"])
            if r["kpis"]:
                empresa_data["reportes"]["control_interno"] = r
                print(f"  Control {len(r['kpis'])} KPIs sem={r['estado']}")

        # ── Productividad
        if scanned.get("productividad_ds"):
            r = build_productividad(scanned["productividad_ds"])
            if r["kpis"]:
                empresa_data["reportes"]["productividad"] = r
                print(f"  Productividad {len(r['kpis'])} KPIs")

        # ── Fill Rate / Avance (planificacion o inventario)
        inv_sc = scanned.get("inventario", {})
        if inv_sc:
            fr = build_fill_rate(inv_sc)
            if fr["kpis"]:
                empresa_data["reportes"]["fill_rate"] = fr
                print(f"  Fill Rate {fr['kpis'][0]['valor'] if fr['kpis'] else '—'}")
            av, avance_pct = build_avance(inv_sc)
            if av["kpis"]:
                empresa_data["reportes"]["margen_variable_pag2"] = av
                print(f"  Avance PPTO {fmt_pct(avance_pct)} sem={av['estado']}")

        summary["empresas"][empresa] = empresa_data

    if holding_ventas: summary["holding_ventas"] = fmt_soles(holding_ventas)
    if holding_mora_vals:
        summary["holding_mora"] = f"{sum(holding_mora_vals)/len(holding_mora_vals):.1f}%"

    out = OUTPUT_DIR / "summaries.json"
    out.write_text(json.dumps(summary, ensure_ascii=False, indent=2))
    print(f"\n✓ summaries.json guardado ({out})")

    hist_path = HIST_DIR / "historico.json"
    try:
        historico = json.loads(hist_path.read_text()) if hist_path.exists() else {
            "generado_por": "JUANITO — Módulo Histórico",
            "ultima_actualizacion": HOY,
            "empresas": {"PAUNO": {"meses": []}}
        }
        mes_label = datetime.date.today().strftime("%b-%y")
        for emp, ed in summary.get("empresas", {}).items():
            meses = historico["empresas"].setdefault(emp, {"meses": []})["meses"]
            ventas_kpi = next((k["valor"] for k in ed.get("reportes",{}).get("margen_variable",{}).get("kpis",[]) if "venta" in k.get("label","").lower()), "—")
            mora_kpi   = next((k["valor"] for k in ed.get("reportes",{}).get("cuentas_por_cobrar",{}).get("kpis",[]) if "mora" in k.get("label","").lower()), "—")
            entrada = {"mes": mes_label, "ventas": ventas_kpi, "mora": mora_kpi}
            if not meses or meses[-1]["mes"] != mes_label:
                meses.append(entrada)
                if len(meses) > 12: meses.pop(0)
        historico["ultima_actualizacion"] = HOY
        hist_path.write_text(json.dumps(historico, ensure_ascii=False, indent=2))
        print("✓ historico.json actualizado")
    except Exception as e:
        print(f"  Histórico error: {e}")

    print("\n=== COMPLETADO ===")

if __name__ == "__main__":
    main()
