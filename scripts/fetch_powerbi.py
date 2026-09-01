#!/usr/bin/env python3
"""
JUANITO — Power BI Data Fetcher (v2)
Autentica via ROPC, descubre medidas, ejecuta DAX y genera summaries.json
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

DATASET_IDS = {
    "PAUNO": {
        "cxc":        "2eec70cd-0820-408f-b938-a2cd547b0c18",
        "cxp":        "45a8ab8d-e162-4398-a260-3a9a5f90829f",
        "margen":     "38076daa-d2cd-4a93-858a-82c0a4cf8cb6",
        "mermas":     "35866214-f4da-45a3-a5a2-aa0c8caffe78",
        "compras":    "06408938-8202-424e-80c0-b42c178dabde",
        "inventario": "0e27d784-41a4-48f0-9208-60210119f0a7",
    }
}

def get_token():
    r = requests.post(TOKEN_URL, data={
        "grant_type": "password", "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET, "username": USERNAME,
        "password": PASSWORD, "scope": PBI_SCOPE,
    }, timeout=30)
    r.raise_for_status()
    return r.json()["access_token"]

def dax(token, ws_id, dataset_id, query, label=""):
    r = requests.post(
        f"{PBI_BASE}/groups/{ws_id}/datasets/{dataset_id}/executeQueries",
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json={"queries": [{"query": query}], "serializerSettings": {"includeNulls": True}},
        timeout=45,
    )
    if not r.ok:
        print(f"    DAX error {label} {r.status_code}: {r.text[:150]}")
        return []
    return r.json().get("results", [{}])[0].get("tables", [{}])[0].get("rows", [])

def dax_row(token, ws_id, dataset_id, query, label=""):
    rows = dax(token, ws_id, dataset_id, query, label)
    return rows[0] if rows else {}

def discover_measures(token, ws_id, dataset_id):
    """Retorna set de nombres de medidas disponibles en el dataset."""
    rows = dax(token, ws_id, dataset_id, "EVALUATE INFO.MEASURES()", "discover")
    measures = set()
    for r in rows:
        name = r.get("[Name]") or r.get("Name") or ""
        if name:
            measures.add(name.strip())
    return measures

def pick(row, *keys):
    """Extrae valor de una fila DAX probando múltiples nombres de columna."""
    for k in keys:
        v = row.get(f"[{k}]") or row.get(k)
        if v is not None:
            return v
    return None

def fmt_soles(v, decimals=0):
    if v is None: return "—"
    try:
        n = float(str(v).replace(",", ".").replace("%", "").replace("S/", "").replace(" ", ""))
        if abs(n) >= 1_000_000:
            return f"S/{n/1_000_000:.2f}M"
        elif abs(n) >= 1_000:
            return f"S/{n:,.0f}"
        return f"S/{n:.{decimals}f}"
    except: return str(v)

def fmt_pct(v):
    if v is None: return "—"
    try:
        n = float(str(v).replace(",", ".").replace("%", ""))
        if abs(n) < 1: n *= 100
        return f"{n:.1f}%"
    except: return str(v)

def sem_by_pct(v, red_above=None, yellow_above=None, red_below=None, yellow_below=None):
    if v is None: return "green"
    try:
        n = float(str(v).replace(",", ".").replace("%", ""))
        if abs(n) < 1: n *= 100
        if red_above and n >= red_above: return "red"
        if yellow_above and n >= yellow_above: return "yellow"
        if red_below and n <= red_below: return "red"
        if yellow_below and n <= yellow_below: return "yellow"
        return "green"
    except: return "green"

# ──────────────────────────────────────────────────────────────────────────────
# Query functions — intentan múltiples nombres de medida
# ──────────────────────────────────────────────────────────────────────────────

def query_cxc(token, ws_id, dataset_id, measures, empresa):
    """CxC — Morosidad, tramos, clientes críticos."""
    # Morosidad
    mora_names = [m for m in measures if "morosidad" in m.lower() or "mora" in m.lower()]
    vencer_names = [m for m in measures if "vencer" in m.lower()]
    total_names = [m for m in measures if "total" in m.lower() and ("cxc" in m.lower() or "cobrar" in m.lower())]

    measures_str = ", ".join(
        [f'"{m}", [{m}]' for m in (mora_names[:1] or ["% Morosidad"]) + (vencer_names[:1] or ["Por Vencer"]) + (total_names[:1] or [])]
    )
    if not measures_str:
        measures_str = '"mora", [% Morosidad], "vencer", [Por Vencer]'

    q = f"EVALUATE ROW({measures_str})"
    row = dax_row(token, ws_id, dataset_id, q, "cxc")

    mora_raw = pick(row, *(mora_names[:1] or ["% Morosidad"]), "% Morosidad", "Morosidad", "Mora")
    vencer_raw = pick(row, *(vencer_names[:1] or ["Por Vencer"]), "Por Vencer", "Vencer")

    mora_pct = None
    if mora_raw is not None:
        try:
            mora_pct = float(mora_raw)
            if abs(mora_pct) < 1: mora_pct *= 100
        except: pass

    sem = sem_by_pct(mora_pct, red_above=20, yellow_above=15)
    razon = f"Morosidad {fmt_pct(mora_pct)}" + (" — crítica" if sem=="red" else " — sobre meta" if sem=="yellow" else " — OK")

    kpis = []
    if mora_pct is not None:
        kpis.append({"label": "Morosidad", "valor": fmt_pct(mora_pct), "meta": "15%", "estado": sem})
    if vencer_raw is not None:
        kpis.append({"label": "CxC por vencer", "valor": fmt_soles(vencer_raw)})

    # Intentar tramos (0-15d, 16-30d, +30d)
    tramos_measures = {
        "0-15d": [m for m in measures if ("0" in m and "15" in m) or "15 d" in m.lower()],
        "16-30d": [m for m in measures if ("16" in m or "30" in m) and "d" in m.lower()],
        "+30d": [m for m in measures if "+30" in m or "mayor" in m.lower() and "30" in m],
    }
    tramos = []
    for label, ms in tramos_measures.items():
        if ms:
            r2 = dax_row(token, ws_id, dataset_id, f"EVALUATE ROW(\"{label}\", [{ms[0]}])", f"cxc_tramo_{label}")
            val = pick(r2, label, ms[0])
            if val is not None:
                tramos.append({"label": label, "valor": fmt_soles(val)})

    result = {"estado": sem, "alerta": razon if sem != "green" else None, "kpis": kpis}
    if tramos:
        result["tramos"] = tramos
    return result, sem, razon, mora_pct, vencer_raw

def query_cxp(token, ws_id, dataset_id, measures):
    """CxP — Días, total, proveedores."""
    dias_names = [m for m in measures if "dia" in m.lower() and ("cxp" in m.lower() or "pagar" in m.lower() or "plazo" in m.lower())]
    total_names = [m for m in measures if "total" in m.lower() and ("cxp" in m.lower() or "pagar" in m.lower())]

    parts = []
    for ms, fallback in [(dias_names, "Dias CxP"), (total_names, "CxP Total")]:
        name = ms[0] if ms else fallback
        parts.append(f'"{name}", [{name}]')

    q = f"EVALUATE ROW({', '.join(parts)})"
    row = dax_row(token, ws_id, dataset_id, q, "cxp")

    dias_raw = pick(row, *(dias_names[:1] or []), "Dias CxP", "Días CxP", "DiasCxP", "Plazo Pago")
    total_raw = pick(row, *(total_names[:1] or []), "CxP Total", "Total CxP", "Total Pagar")

    dias = None
    if dias_raw is not None:
        try: dias = float(dias_raw)
        except: pass

    sem = "green"
    if dias:
        if dias > 90: sem = "red"
        elif dias > 60: sem = "yellow"

    kpis = []
    if dias is not None:
        kpis.append({"label": "Días CxP", "valor": f"{dias:.0f} días", "meta": "<90d", "estado": sem})
    if total_raw is not None:
        kpis.append({"label": "CxP Total", "valor": fmt_soles(total_raw)})

    return {"estado": sem, "alerta": f"CxP {dias:.0f}d — revisar flujo" if sem != "green" and dias else None, "kpis": kpis}

def query_margen(token, ws_id, dataset_id, measures):
    """Margen Variable y Ventas."""
    ventas_names  = [m for m in measures if "venta" in m.lower() and ("actual" in m.lower() or "mes" in m.lower())]
    margen_names  = [m for m in measures if "margen" in m.lower() and "variable" in m.lower()]
    precio_names  = [m for m in measures if "precio" in m.lower() and "kg" in m.lower()]

    parts = []
    for ms, fallback in [(ventas_names, "Ventas Mes Actual"), (margen_names, "Margen Variable"), (precio_names, "Precio/kg")]:
        name = ms[0] if ms else fallback
        parts.append(f'"{name}", [{name}]')

    row = dax_row(token, ws_id, dataset_id, f"EVALUATE ROW({', '.join(parts)})", "margen")

    ventas_raw = pick(row, *(ventas_names[:1] or []), "Ventas Mes Actual", "Ventas Actuales", "Ventas")
    margen_raw = pick(row, *(margen_names[:1] or []), "Margen Variable", "Margen", "% Margen Variable")
    precio_raw = pick(row, *(precio_names[:1] or []), "Precio/kg", "Precio kg")

    margen_pct = None
    if margen_raw is not None:
        try:
            margen_pct = float(margen_raw)
            if abs(margen_pct) < 1: margen_pct *= 100
        except: pass

    sem = sem_by_pct(margen_pct, red_below=46, yellow_below=52)

    kpis = []
    if ventas_raw is not None:
        kpis.append({"label": "Ventas mes", "valor": fmt_soles(ventas_raw)})
    if margen_pct is not None:
        kpis.append({"label": "Margen variable", "valor": fmt_pct(margen_pct), "meta": "52%", "estado": sem})
    if precio_raw is not None:
        kpis.append({"label": "Precio/kg", "valor": f"S/{float(precio_raw):.2f}" if precio_raw else "—"})

    return {
        "estado": sem,
        "alerta": f"Margen {fmt_pct(margen_pct)} — bajo meta 52%" if sem != "green" and margen_pct else None,
        "kpis": kpis,
        "_ventas": ventas_raw, "_margen": margen_pct
    }

def query_mermas(token, ws_id, dataset_id, measures):
    """Mermas por planta."""
    ate_names  = [m for m in measures if "merma" in m.lower() and "ate" in m.lower()]
    pach_names = [m for m in measures if "merma" in m.lower() and ("pach" in m.lower() or "lurin" in m.lower())]
    total_names = [m for m in measures if "merma" in m.lower() and ("total" in m.lower() or "general" in m.lower())]
    terc_names = [m for m in measures if "merma" in m.lower() and "tercer" in m.lower()]

    # Fallback query with common names
    q_parts = []
    for ms, fb in [(ate_names, "% Merma Ate"), (pach_names, "% Merma Pachacamac"), (total_names, "% Merma Total"), (terc_names, "% Merma Terceros")]:
        n = ms[0] if ms else fb
        q_parts.append(f'"{n}", [{n}]')

    row = dax_row(token, ws_id, dataset_id, f"EVALUATE ROW({', '.join(q_parts)})", "mermas")

    ate_raw  = pick(row, *(ate_names[:1] or []), "% Merma Ate", "Merma Ate", "% Mermas Ate")
    pach_raw = pick(row, *(pach_names[:1] or []), "% Merma Pachacamac", "Merma Pachacamac", "% Merma Lurin")
    tot_raw  = pick(row, *(total_names[:1] or []), "% Merma Total", "Merma Total")
    terc_raw = pick(row, *(terc_names[:1] or []), "% Merma Terceros", "Merma Terceros")

    def to_pct(v):
        if v is None: return None
        try:
            n = float(v)
            if abs(n) < 1: n *= 100
            return n
        except: return None

    ate  = to_pct(ate_raw)
    pach = to_pct(pach_raw)
    tot  = to_pct(tot_raw)
    terc = to_pct(terc_raw)

    worst = max(filter(None, [ate, pach, tot, terc]), default=None)
    sem = "green"
    if worst:
        if worst > 3: sem = "red"
        elif worst > 2: sem = "yellow"

    kpis = []
    for label, val, meta in [("Merma Ate", ate, "2%"), ("Merma Pachacamac", pach, "2%"), ("Merma Total", tot, "2%"), ("Merma Terceros", terc, "2%")]:
        if val is not None:
            s = "green"
            if val > 3: s = "red"
            elif val > 2: s = "yellow"
            kpis.append({"label": label, "valor": f"{val:.2f}%", "meta": meta, "estado": s})

    return {
        "estado": sem,
        "alerta": f"Merma {worst:.2f}% — sobre meta 2%" if sem != "green" and worst else None,
        "kpis": kpis
    }

def query_compras(token, ws_id, dataset_id, measures):
    """Ratio Compras = Consumo / Compra."""
    ratio_names   = [m for m in measures if "ratio" in m.lower() and ("compra" in m.lower() or "consumo" in m.lower())]
    consumo_names = [m for m in measures if "consumo" in m.lower() and "compra" not in m.lower()]
    compra_names  = [m for m in measures if "compra" in m.lower() and "ratio" not in m.lower() and "consumo" not in m.lower()]

    parts = []
    for ms, fb in [(ratio_names, "Ratio Compras"), (consumo_names, "Consumo"), (compra_names, "Compras")]:
        n = ms[0] if ms else fb
        parts.append(f'"{n}", [{n}]')

    row = dax_row(token, ws_id, dataset_id, f"EVALUATE ROW({', '.join(parts)})", "compras")

    ratio_raw   = pick(row, *(ratio_names[:1] or []), "Ratio Compras", "Ratio Consumo/Compra", "% Ratio")
    consumo_raw = pick(row, *(consumo_names[:1] or []), "Consumo", "Total Consumo")
    compra_raw  = pick(row, *(compra_names[:1] or []), "Compras", "Total Compras")

    ratio = None
    if ratio_raw is not None:
        try:
            ratio = float(ratio_raw)
            if abs(ratio) < 2: ratio *= 100
        except: pass
    elif consumo_raw and compra_raw:
        try:
            c, p = float(consumo_raw), float(compra_raw)
            if p > 0: ratio = (c / p) * 100
        except: pass

    sem = "green"
    if ratio is not None:
        if ratio > 130 or ratio < 70: sem = "red"
        elif ratio > 110 or ratio < 80: sem = "yellow"

    kpis = []
    if ratio is not None:
        interp = "jala inventario" if ratio > 100 else "sobre-compra" if ratio < 80 else "normal"
        kpis.append({"label": "Ratio Consumo/Compra", "valor": f"{ratio:.1f}%", "meta": "80-100%", "estado": sem, "interpretacion": interp})
    if consumo_raw is not None:
        kpis.append({"label": "Consumo", "valor": fmt_soles(consumo_raw)})
    if compra_raw is not None:
        kpis.append({"label": "Compras", "valor": fmt_soles(compra_raw)})

    return {
        "estado": sem,
        "alerta": f"Ratio {ratio:.1f}% — {'jala inventario' if ratio and ratio>100 else 'sobre-compra'}" if sem != "green" and ratio else None,
        "kpis": kpis, "_ratio": ratio
    }

def query_inventario(token, ws_id, dataset_id, measures):
    """S&OP / Inventario / Dead Stock."""
    dead_names  = [m for m in measures if "dead" in m.lower() or ("muerto" in m.lower() and "stock" in m.lower())]
    deadpct_names = [m for m in measures if ("%" in m or "pct" in m.lower() or "porcentaje" in m.lower()) and ("dead" in m.lower() or "inmovilizado" in m.lower())]
    total_names = [m for m in measures if "inventario" in m.lower() and "total" in m.lower()]
    working_names = [m for m in measures if "working" in m.lower() or "activo" in m.lower()]

    parts = []
    for ms, fb in [(dead_names, "Dead Stock"), (deadpct_names, "% Dead Stock"), (total_names, "Inventario Total"), (working_names, "Working Stock")]:
        n = ms[0] if ms else fb
        parts.append(f'"{n}", [{n}]')

    row = dax_row(token, ws_id, dataset_id, f"EVALUATE ROW({', '.join(parts)})", "inventario")

    dead_raw    = pick(row, *(dead_names[:1] or []), "Dead Stock", "Stock Muerto", "Inmovilizado")
    deadpct_raw = pick(row, *(deadpct_names[:1] or []), "% Dead Stock", "% Inmovilizado", "Pct Dead")
    total_raw   = pick(row, *(total_names[:1] or []), "Inventario Total", "Total Inventario")
    working_raw = pick(row, *(working_names[:1] or []), "Working Stock", "Stock Activo")

    dead_pct = None
    if deadpct_raw is not None:
        try:
            dead_pct = float(deadpct_raw)
            if abs(dead_pct) < 1: dead_pct *= 100
        except: pass
    elif dead_raw and total_raw:
        try:
            dead_pct = float(dead_raw) / float(total_raw) * 100
        except: pass

    sem = sem_by_pct(dead_pct, red_above=10, yellow_above=5)

    kpis = []
    if dead_raw is not None:
        kpis.append({"label": "Dead Stock", "valor": fmt_soles(dead_raw), "meta": "<S/500K", "estado": sem})
    if dead_pct is not None:
        kpis.append({"label": "% Dead Stock", "valor": f"{dead_pct:.1f}%", "meta": "<5%", "estado": sem})
    if total_raw is not None:
        kpis.append({"label": "Inventario Total", "valor": fmt_soles(total_raw)})
    if working_raw is not None:
        kpis.append({"label": "Working Stock", "valor": fmt_soles(working_raw)})

    return {
        "estado": sem,
        "alerta": f"Dead Stock {dead_pct:.1f}% del inventario" if sem != "green" and dead_pct else None,
        "kpis": kpis
    }

def query_productividad(token, ws_id, dataset_id, measures):
    """Productividad planilla/kg."""
    planilla_kg_names = [m for m in measures if ("planilla" in m.lower() or "productividad" in m.lower()) and "kg" in m.lower()]
    planilla_names    = [m for m in measures if "planilla" in m.lower() and "kg" not in m.lower()]
    kg_names          = [m for m in measures if "kg" in m.lower() and "produccion" in m.lower()]

    parts = []
    for ms, fb in [(planilla_kg_names, "Planilla/kg"), (planilla_names, "Planilla Total"), (kg_names, "Kg Producidos")]:
        n = ms[0] if ms else fb
        parts.append(f'"{n}", [{n}]')

    row = dax_row(token, ws_id, dataset_id, f"EVALUATE ROW({', '.join(parts)})", "productividad")

    pkg_raw     = pick(row, *(planilla_kg_names[:1] or []), "Planilla/kg", "S//kg", "Productividad")
    planilla_raw = pick(row, *(planilla_names[:1] or []), "Planilla Total", "Costo Planilla")
    kg_raw      = pick(row, *(kg_names[:1] or []), "Kg Producidos", "Produccion Kg")

    kpis = []
    if pkg_raw is not None:
        kpis.append({"label": "S/kg producido", "valor": f"S/{float(pkg_raw):.2f}" if pkg_raw else "—"})
    if planilla_raw is not None:
        kpis.append({"label": "Planilla total", "valor": fmt_soles(planilla_raw)})
    if kg_raw is not None:
        kpis.append({"label": "Kg producidos", "valor": f"{float(kg_raw):,.0f} kg" if kg_raw else "—"})

    return {"estado": "green", "alerta": None, "kpis": kpis}

def query_fill_rate(token, ws_id, dataset_id, measures):
    """Fill Rate."""
    fill_names = [m for m in measures if "fill" in m.lower() or "tasa" in m.lower() and "atenci" in m.lower()]

    parts = []
    for ms, fb in [(fill_names, "Fill Rate")]:
        n = ms[0] if ms else fb
        parts.append(f'"{n}", [{n}]')

    row = dax_row(token, ws_id, dataset_id, f"EVALUATE ROW({', '.join(parts)})", "fill_rate")

    fill_raw = pick(row, *(fill_names[:1] or []), "Fill Rate", "% Fill Rate", "Tasa Atencion")

    fill_pct = None
    if fill_raw is not None:
        try:
            fill_pct = float(fill_raw)
            if abs(fill_pct) < 1: fill_pct *= 100
        except: pass

    sem = sem_by_pct(fill_pct, red_below=85, yellow_below=95)

    kpis = []
    if fill_pct is not None:
        kpis.append({"label": "Fill Rate", "valor": fmt_pct(fill_pct), "meta": "98%", "estado": sem})
        kpis.append({"label": "Brecha", "valor": f"-{98 - fill_pct:.1f}pp" if fill_pct < 98 else "0pp"})

    return {
        "estado": sem,
        "alerta": f"Fill Rate {fmt_pct(fill_pct)} — bajo meta 98%" if sem != "green" and fill_pct else None,
        "kpis": kpis
    }

def query_control_interno(token, ws_id, dataset_id, measures):
    """Control Interno — Cumplimiento."""
    cumpl_names = [m for m in measures if "cumplimiento" in m.lower() or "conformidad" in m.lower()]
    criticos_names = [m for m in measures if "critico" in m.lower() and ("punto" in m.lower() or "hallazgo" in m.lower())]

    parts = []
    for ms, fb in [(cumpl_names, "% Cumplimiento"), (criticos_names, "Puntos Criticos")]:
        n = ms[0] if ms else fb
        parts.append(f'"{n}", [{n}]')

    row = dax_row(token, ws_id, dataset_id, f"EVALUATE ROW({', '.join(parts)})", "control_interno")

    cumpl_raw   = pick(row, *(cumpl_names[:1] or []), "% Cumplimiento", "Cumplimiento", "Conformidad")
    criticos_raw = pick(row, *(criticos_names[:1] or []), "Puntos Criticos", "Hallazgos Criticos")

    cumpl_pct = None
    if cumpl_raw is not None:
        try:
            cumpl_pct = float(cumpl_raw)
            if abs(cumpl_pct) < 1: cumpl_pct *= 100
        except: pass

    sem = sem_by_pct(cumpl_pct, red_below=70, yellow_below=85)

    kpis = []
    if cumpl_pct is not None:
        kpis.append({"label": "Cumplimiento", "valor": fmt_pct(cumpl_pct), "meta": "85%", "estado": sem})
    if criticos_raw is not None:
        kpis.append({"label": "Puntos críticos", "valor": str(criticos_raw), "meta": "0", "estado": "red" if float(criticos_raw or 0) > 0 else "green"})

    return {
        "estado": sem,
        "alerta": f"Cumplimiento {fmt_pct(cumpl_pct)} — bajo meta 85%" if sem != "green" and cumpl_pct else None,
        "kpis": kpis
    }

def query_avance_presupuesto(token, ws_id, dataset_id, measures):
    """Avance de Ventas vs Presupuesto."""
    avance_names = [m for m in measures if ("avance" in m.lower() or "cumpl" in m.lower()) and ("ppto" in m.lower() or "presupuesto" in m.lower() or "budget" in m.lower())]
    real_names   = [m for m in measures if "venta" in m.lower() and ("real" in m.lower() or "actual" in m.lower())]
    ppto_names   = [m for m in measures if "ppto" in m.lower() or "presupuesto" in m.lower() or "budget" in m.lower()]

    parts = []
    for ms, fb in [(avance_names, "% Avance Presupuesto"), (real_names, "Ventas Reales"), (ppto_names, "Presupuesto")]:
        n = ms[0] if ms else fb
        parts.append(f'"{n}", [{n}]')

    row = dax_row(token, ws_id, dataset_id, f"EVALUATE ROW({', '.join(parts)})", "avance_ppto")

    avance_raw = pick(row, *(avance_names[:1] or []), "% Avance Presupuesto", "Avance PPTO", "% PPTO")
    real_raw   = pick(row, *(real_names[:1] or []), "Ventas Reales", "Ventas Actuales")
    ppto_raw   = pick(row, *(ppto_names[:1] or []), "Presupuesto", "PPTO")

    avance_pct = None
    if avance_raw is not None:
        try:
            avance_pct = float(avance_raw)
            if abs(avance_pct) < 1: avance_pct *= 100
        except: pass
    elif real_raw and ppto_raw:
        try:
            avance_pct = float(real_raw) / float(ppto_raw) * 100
        except: pass

    sem = sem_by_pct(avance_pct, red_below=60, yellow_below=80)

    kpis = []
    if avance_pct is not None:
        kpis.append({"label": "Avance vs Ppto", "valor": fmt_pct(avance_pct), "meta": "100%", "estado": sem})
    if real_raw is not None:
        kpis.append({"label": "Ventas reales", "valor": fmt_soles(real_raw)})
    if ppto_raw is not None:
        kpis.append({"label": "Presupuesto", "valor": fmt_soles(ppto_raw)})

    return {
        "estado": sem,
        "alerta": f"Avance {fmt_pct(avance_pct)} vs presupuesto" if sem != "green" and avance_pct else None,
        "kpis": kpis, "_avance": avance_pct
    }

# ──────────────────────────────────────────────────────────────────────────────
# Main
# ──────────────────────────────────────────────────────────────────────────────

def main():
    print("=== JUANITO Power BI Fetcher v2 ===")
    print(f"Fecha: {HOY}")

    print("\n[1] Autenticando...")
    try:
        token = get_token()
        print("    Token OK")
    except Exception as e:
        print(f"    ERROR auth: {e}"); sys.exit(1)

    summary = {
        "fecha": datetime.date.today().strftime("%d/%m/%Y"),
        "fecha_actualizacion": HOY,
        "generado_por": "JUANITO — Power BI Direct v2",
        "semaforos": {}, "semaforo_razon": {},
        "holding_ventas": "—", "holding_mora": "—",
        "agenda_ceo": {"decidir_hoy": [], "escalar_semana": [], "monitorear": []},
        "empresas": {}
    }

    holding_ventas = 0
    holding_mora_vals = []

    for empresa, ws_id in WORKSPACES.items():
        print(f"\n[{empresa}] Descubriendo medidas...")
        ids = DATASET_IDS.get(empresa, {})
        empresa_data = {"reportes": {}}

        # Descubrir medidas de cada dataset
        dataset_measures = {}
        for key, did in ids.items():
            ms = discover_measures(token, ws_id, did)
            dataset_measures[key] = ms
            print(f"  {key}: {len(ms)} medidas encontradas")
            if ms:
                sample = sorted(list(ms))[:8]
                print(f"    Muestra: {sample}")

        # ── CxC ──────────────────────────────────────────────────────────────
        if "cxc" in ids:
            print(f"  Consultando CxC...")
            cxc_result, sem, razon, mora_pct, por_vencer = query_cxc(
                token, ws_id, ids["cxc"], dataset_measures.get("cxc", set()), empresa
            )
            empresa_data["reportes"]["cuentas_por_cobrar"] = cxc_result
            summary["semaforos"][empresa] = sem
            summary["semaforo_razon"][empresa] = razon
            if mora_pct: holding_mora_vals.append(mora_pct)
            print(f"    mora={fmt_pct(mora_pct)} sem={sem}")
            if sem == "red":
                summary["agenda_ceo"]["decidir_hoy"].append({
                    "empresa": empresa,
                    "texto": f"Mora {fmt_pct(mora_pct)} en {empresa} — vencido {fmt_soles(por_vencer)}. Activar cobranza intensiva.",
                    "responsable": "Gerencia Financiera"
                })

        # ── CxP ──────────────────────────────────────────────────────────────
        if "cxp" in ids:
            print(f"  Consultando CxP...")
            cxp_result = query_cxp(token, ws_id, ids["cxp"], dataset_measures.get("cxp", set()))
            empresa_data["reportes"]["cuentas_por_pagar"] = cxp_result
            dias_val = next((k["valor"] for k in cxp_result.get("kpis", []) if "Días" in k.get("label", "")), None)
            print(f"    {dias_val or 'sin datos'} sem={cxp_result.get('estado')}")
            if cxp_result.get("estado") == "red":
                summary["agenda_ceo"]["escalar_semana"].append({
                    "empresa": empresa,
                    "texto": f"CxP {empresa} {dias_val} — riesgo con proveedores.",
                    "responsable": "Finanzas"
                })

        # ── Margen ────────────────────────────────────────────────────────────
        if "margen" in ids:
            print(f"  Consultando Margen...")
            mv = query_margen(token, ws_id, ids["margen"], dataset_measures.get("margen", set()))
            empresa_data["reportes"]["margen_variable"] = {k: v for k, v in mv.items() if not k.startswith("_")}
            print(f"    margen={fmt_pct(mv.get('_margen'))} ventas={fmt_soles(mv.get('_ventas'))}")
            if mv.get("_ventas"):
                try: holding_ventas += float(mv["_ventas"])
                except: pass
            if mv.get("estado") == "red":
                summary["agenda_ceo"]["escalar_semana"].append({
                    "empresa": empresa,
                    "texto": f"Margen variable {fmt_pct(mv.get('_margen'))} en {empresa} — bajo meta 52%.",
                    "responsable": "Comercial / Costos"
                })

        # ── Mermas ────────────────────────────────────────────────────────────
        if "mermas" in ids:
            print(f"  Consultando Mermas...")
            mermas = query_mermas(token, ws_id, ids["mermas"], dataset_measures.get("mermas", set()))
            empresa_data["reportes"]["mermas"] = mermas
            print(f"    {len(mermas.get('kpis', []))} KPIs sem={mermas.get('estado')}")

        # ── Compras ───────────────────────────────────────────────────────────
        if "compras" in ids:
            print(f"  Consultando Compras...")
            compras = query_compras(token, ws_id, ids["compras"], dataset_measures.get("compras", set()))
            empresa_data["reportes"]["compras"] = {k: v for k, v in compras.items() if not k.startswith("_")}
            print(f"    ratio={compras.get('_ratio', '—')} sem={compras.get('estado')}")
            if compras.get("estado") == "red":
                summary["agenda_ceo"]["decidir_hoy"].append({
                    "empresa": empresa,
                    "texto": f"Ratio Compras {compras.get('_ratio', '?'):.1f}% — {'jalando inventario, pedir compra urgente' if (compras.get('_ratio') or 0) > 100 else 'sobre-compra, frenar órdenes'}.",
                    "responsable": "Logística"
                })

        # ── Inventario / S&OP ─────────────────────────────────────────────────
        if "inventario" in ids:
            print(f"  Consultando Inventario/S&OP...")
            inv = query_inventario(token, ws_id, ids["inventario"], dataset_measures.get("inventario", set()))
            empresa_data["reportes"]["sop_inventario"] = inv
            print(f"    {len(inv.get('kpis', []))} KPIs sem={inv.get('estado')}")

            # Reutilizar dataset de inventario para productividad, fill rate, control, avance
            ms_inv = dataset_measures.get("inventario", set())

            print(f"  Consultando Productividad (mismo dataset)...")
            prod = query_productividad(token, ws_id, ids["inventario"], ms_inv)
            if prod.get("kpis"):
                empresa_data["reportes"]["productividad"] = prod

            print(f"  Consultando Fill Rate (mismo dataset)...")
            fr = query_fill_rate(token, ws_id, ids["inventario"], ms_inv)
            if fr.get("kpis"):
                empresa_data["reportes"]["fill_rate"] = fr

            print(f"  Consultando Control Interno (mismo dataset)...")
            ci = query_control_interno(token, ws_id, ids["inventario"], ms_inv)
            if ci.get("kpis"):
                empresa_data["reportes"]["control_interno"] = ci

            print(f"  Consultando Avance Presupuesto (dataset margen)...")

        # Avance presupuesto — usar dataset margen o inventario
        for ds_key in ["margen", "inventario"]:
            if ds_key in ids:
                av = query_avance_presupuesto(token, ws_id, ids[ds_key], dataset_measures.get(ds_key, set()))
                if av.get("kpis"):
                    empresa_data["reportes"]["margen_variable_pag2"] = {k: v for k, v in av.items() if not k.startswith("_")}
                    print(f"    avance={fmt_pct(av.get('_avance'))} sem={av.get('estado')}")
                    if av.get("estado") in ("red", "yellow"):
                        summary["agenda_ceo"]["monitorear"].append({
                            "empresa": empresa,
                            "texto": f"Avance presupuesto ventas {fmt_pct(av.get('_avance'))} en {empresa}.",
                        })
                    break

        summary["empresas"][empresa] = empresa_data

    # Holding totals
    if holding_ventas: summary["holding_ventas"] = fmt_soles(holding_ventas)
    if holding_mora_vals:
        summary["holding_mora"] = fmt_pct(sum(holding_mora_vals) / len(holding_mora_vals))

    # Guardar
    out = OUTPUT_DIR / "summaries.json"
    out.write_text(json.dumps(summary, ensure_ascii=False, indent=2))
    print(f"\n✓ summaries.json guardado ({out})")

    # Histórico
    hist_path = HIST_DIR / "historico.json"
    try:
        historico = json.loads(hist_path.read_text()) if hist_path.exists() else {
            "generado_por": "JUANITO — Módulo Histórico",
            "ultima_actualizacion": HOY,
            "empresas": {"PAUNO": {"meses": []}, "AMAUTA": {"meses": []}, "NEOPACK": {"meses": []}}
        }
        mes_label = datetime.date.today().strftime("%b-%y")
        for emp, ed in summary.get("empresas", {}).items():
            meses = historico["empresas"].setdefault(emp, {"meses": []})["meses"]
            ventas_kpi = next((k["valor"] for k in ed.get("reportes", {}).get("margen_variable", {}).get("kpis", []) if "venta" in k.get("label", "").lower()), "—")
            mora_kpi   = next((k["valor"] for k in ed.get("reportes", {}).get("cuentas_por_cobrar", {}).get("kpis", []) if "mora" in k.get("label", "").lower()), "—")
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
