#!/usr/bin/env python3
"""
JUANITO — Power BI Data Fetcher
Autentica via ROPC, ejecuta DAX queries y genera JSON para JUANITO.
Corre automáticamente via GitHub Actions cada mañana a las 7:30am Lima.
"""

import os, json, requests, datetime, sys
from pathlib import Path

warnings_filter = True
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

# ── Workspaces y Datasets ─────────────────────────────────────────────────────
WORKSPACES = {
    "PAUNO":  "461932ad-b5ec-4fd6-aa97-f1fc7bdc5169",
    "AMAUTA": "8266b260-4af1-41fa-a011-cff28e984a6d",
    "NEOPACK":"317a3742-3a89-4a9b-b9ab-3a5ee8be2505",
}

# Dataset IDs conocidos (PAUNO confirmados; AMAUTA/NEOPACK se descubren en runtime)
DATASET_IDS = {
    "PAUNO": {
        "cxc":     "2eec70cd-0820-408f-b938-a2cd547b0c18",
        "cxp":     "45a8ab8d-e162-4398-a260-3a9a5f90829f",
        "margen":  "38076daa-d2cd-4a93-858a-82c0a4cf8cb6",
        "mermas":  "35866214-f4da-45a3-a5a2-aa0c8caffe78",
        "compras": "06408938-8202-424e-80c0-b42c178dabde",
        "inventario": "0e27d784-41a4-48f0-9208-60210119f0a7",
    }
}

# ── Auth ──────────────────────────────────────────────────────────────────────
def get_token():
    r = requests.post(TOKEN_URL, data={
        "grant_type": "password", "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET, "username": USERNAME,
        "password": PASSWORD, "scope": PBI_SCOPE,
    }, timeout=30)
    r.raise_for_status()
    return r.json()["access_token"]

# ── DAX helper ────────────────────────────────────────────────────────────────
def dax(token, ws_id, dataset_id, query):
    r = requests.post(
        f"{PBI_BASE}/groups/{ws_id}/datasets/{dataset_id}/executeQueries",
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json={"queries": [{"query": query}], "serializerSettings": {"includeNulls": True}},
        timeout=30,
    )
    if not r.ok:
        print(f"    DAX error {r.status_code}: {r.text[:200]}")
        return {}
    rows = r.json().get("results", [{}])[0].get("tables", [{}])[0].get("rows", [])
    return rows[0] if rows else {}

def discover_datasets(token, ws_id):
    """Descubre datasets en un workspace y retorna dict nombre->id."""
    r = requests.get(f"{PBI_BASE}/groups/{ws_id}/datasets",
                     headers={"Authorization": f"Bearer {token}"}, timeout=20)
    if not r.ok: return {}
    return {d["name"]: d["id"] for d in r.json().get("value", [])}

def fmt_soles(v, decimals=0):
    if v is None: return "—"
    try:
        n = float(str(v).replace(",", ".").replace("%", ""))
        return f"S/{n:,.{decimals}f}"
    except: return str(v)

def fmt_pct(v):
    if v is None: return "—"
    try:
        n = float(str(v).replace(",", ".").replace("%", ""))
        if abs(n) < 1: n *= 100
        return f"{n:.1f}%"
    except: return str(v)

# ── Query CxC ─────────────────────────────────────────────────────────────────
def query_cxc(token, ws_id, dataset_id, empresa):
    q = '''EVALUATE ROW(
  "mora_pct", [% Morosidad],
  "por_vencer", [Por Vencer]
)'''
    row = dax(token, ws_id, dataset_id, q)
    mora = row.get("[mora_pct]") or row.get("mora_pct")
    por_vencer = row.get("[por_vencer]") or row.get("por_vencer")

    mora_pct = None
    if mora is not None:
        try:
            mora_pct = float(mora)
            if abs(mora_pct) < 1: mora_pct *= 100
        except: pass

    sem = "green"
    if mora_pct is not None:
        if mora_pct >= 20:   sem = "red"
        elif mora_pct >= 15: sem = "yellow"

    razon = f"Morosidad {fmt_pct(mora_pct)}" + (" — crítica" if sem=="red" else " — sobre meta" if sem=="yellow" else " — OK")

    return {
        "sem": sem, "razon": razon, "mora_pct": mora_pct,
        "por_vencer": por_vencer,
        "kpis_cxc": [
            {"label": "Morosidad", "valor": fmt_pct(mora_pct), "meta": "15%"},
            {"label": "CxC por vencer", "valor": fmt_soles(por_vencer)},
        ]
    }

# ── Query Margen / Ventas ─────────────────────────────────────────────────────
def query_margen(token, ws_id, dataset_id):
    q = '''EVALUATE ROW(
  "ventas_actual", [Ventas Mes Actual],
  "ventas_anterior", [Ventas Mes Anterior],
  "margen", [Margen Variable]
)'''
    row = dax(token, ws_id, dataset_id, q)
    va  = row.get("[ventas_actual]") or row.get("ventas_actual")
    vp  = row.get("[ventas_anterior]") or row.get("ventas_anterior")
    mg  = row.get("[margen]") or row.get("margen")

    kpis = []
    if va is not None:  kpis.append({"label": "Ventas mes actual",  "valor": fmt_soles(va)})
    if vp is not None:  kpis.append({"label": "Ventas mes anterior","valor": fmt_soles(vp)})
    if mg is not None:  kpis.append({"label": "Margen variable",     "valor": fmt_pct(mg)})
    return {"ventas_actual": va, "ventas_anterior": vp, "margen": mg, "kpis": kpis}

# ── Query CxP ─────────────────────────────────────────────────────────────────
def query_cxp(token, ws_id, dataset_id):
    q = 'EVALUATE ROW("dias_cxp", [Dias CxP], "total_cxp", [CxP Total])'
    row = dax(token, ws_id, dataset_id, q)
    dias = row.get("[dias_cxp]") or row.get("dias_cxp")
    total = row.get("[total_cxp]") or row.get("total_cxp")

    sem = "green"
    if dias:
        try:
            d = float(dias)
            if d > 90:   sem = "red"
            elif d > 60: sem = "yellow"
        except: pass

    kpis = []
    if dias:  kpis.append({"label": "Días CxP", "valor": str(dias) + " días", "meta": "<90d"})
    if total: kpis.append({"label": "CxP Total", "valor": fmt_soles(total)})
    return {"sem": sem, "dias": dias, "total": total, "kpis": kpis}

# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    print("=== JUANITO Power BI Fetcher ===")
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
        "generado_por": "JUANITO — Power BI Direct",
        "semaforos": {}, "semaforo_razon": {},
        "holding_ventas": "—", "holding_mora": "—",
        "agenda_ceo": {"decidir_hoy": [], "escalar_semana": []},
        "empresas": {}
    }

    holding_ventas = 0
    holding_mora_vals = []

    for empresa, ws_id in WORKSPACES.items():
        print(f"\n[{empresa}] Consultando...")

        # Descubrir datasets si no están hardcodeados
        if empresa not in DATASET_IDS:
            ds = discover_datasets(token, ws_id)
            DATASET_IDS[empresa] = {}
            for name, did in ds.items():
                nl = name.lower()
                if "cobrar" in nl: DATASET_IDS[empresa]["cxc"] = did
                elif "pagar" in nl: DATASET_IDS[empresa]["cxp"] = did
                elif "margen" in nl or "venta" in nl: DATASET_IDS[empresa]["margen"] = did
                elif "inventario" in nl or "rotacion" in nl: DATASET_IDS[empresa]["inventario"] = did
            print(f"  Datasets descubiertos: {list(DATASET_IDS[empresa].keys())}")

        ids = DATASET_IDS.get(empresa, {})
        empresa_data = {"reportes": {}}

        # CxC
        if "cxc" in ids:
            cxc = query_cxc(token, ws_id, ids["cxc"], empresa)
            print(f"  CxC: mora={cxc['mora_pct']:.1f}% sem={cxc['sem']}" if cxc['mora_pct'] else f"  CxC: {cxc['razon']}")
            empresa_data["reportes"]["cuentas_por_cobrar"] = {
                "estado": cxc["sem"], "alerta": cxc["razon"], "kpis": cxc["kpis_cxc"]
            }
            summary["semaforos"][empresa] = cxc["sem"]
            summary["semaforo_razon"][empresa] = cxc["razon"]
            if cxc["mora_pct"]:
                holding_mora_vals.append(cxc["mora_pct"])
            if cxc["sem"] == "red":
                summary["agenda_ceo"]["decidir_hoy"].append({
                    "empresa": empresa,
                    "texto": f"Mora {fmt_pct(cxc['mora_pct'])} en {empresa} — vencido {fmt_soles(cxc['por_vencer'])}. Activar cobranza intensiva.",
                    "responsable": "Gerencia Financiera"
                })
        else:
            summary["semaforos"][empresa] = "yellow"
            summary["semaforo_razon"][empresa] = "Sin datos de CxC"

        # Margen/Ventas
        if "margen" in ids:
            mv = query_margen(token, ws_id, ids["margen"])
            if mv["kpis"]:
                empresa_data["reportes"]["margen_variable"] = {
                    "estado": "green", "kpis": mv["kpis"]
                }
                empresa_data["reportes"]["ventas"] = {"estado": "green", "kpis": mv["kpis"]}
                if mv["ventas_actual"]:
                    try: holding_ventas += float(mv["ventas_actual"])
                    except: pass

        # CxP
        if "cxp" in ids:
            cxp = query_cxp(token, ws_id, ids["cxp"])
            if cxp["kpis"]:
                empresa_data["reportes"]["cuentas_por_pagar"] = {
                    "estado": cxp["sem"], "kpis": cxp["kpis"]
                }
                if cxp["sem"] == "red" and empresa not in [i["empresa"] for i in summary["agenda_ceo"]["escalar_semana"]]:
                    summary["agenda_ceo"]["escalar_semana"].append({
                        "empresa": empresa,
                        "texto": f"CxP {empresa} en {cxp['dias']} días — riesgo reputacional con proveedores.",
                    })

        summary["empresas"][empresa] = empresa_data

    # Holding totals
    if holding_ventas: summary["holding_ventas"] = fmt_soles(holding_ventas)
    if holding_mora_vals:
        avg = sum(holding_mora_vals) / len(holding_mora_vals)
        summary["holding_mora"] = f"{avg:.1f}%"

    # Guardar
    out = OUTPUT_DIR / "summaries.json"
    out.write_text(json.dumps(summary, ensure_ascii=False, indent=2))
    print(f"\n✓ summaries.json guardado ({out})")

    # Actualizar histórico
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
            ventas_kpi = next((k["valor"] for k in (ed.get("reportes",{}).get("ventas",{}).get("kpis",[]) or []) if "actual" in k.get("label","").lower()), "—")
            mora_kpi   = next((k["valor"] for k in (ed.get("reportes",{}).get("cuentas_por_cobrar",{}).get("kpis",[]) or []) if "mora" in k.get("label","").lower()), "—")
            entrada = {"mes": mes_label, "ventas": ventas_kpi, "mora": mora_kpi}
            if not meses or meses[-1]["mes"] != mes_label:
                meses.append(entrada)
                if len(meses) > 12: meses.pop(0)
        historico["ultima_actualizacion"] = HOY
        hist_path.write_text(json.dumps(historico, ensure_ascii=False, indent=2))
        print(f"✓ historico.json actualizado")
    except Exception as e:
        print(f"  Histórico: {e}")

    print("\n=== COMPLETADO ===")

if __name__ == "__main__":
    main()
