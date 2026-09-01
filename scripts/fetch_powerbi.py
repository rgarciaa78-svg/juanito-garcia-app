#!/usr/bin/env python3
"""
JUANITO — Power BI Data Fetcher
Autentica via ROPC, descubre datasets, ejecuta DAX queries y genera JSON para JUANITO.
"""

import os, json, requests, datetime, sys
from pathlib import Path

TENANT_ID     = os.environ["AZURE_TENANT_ID"]
CLIENT_ID     = os.environ["AZURE_CLIENT_ID"]
CLIENT_SECRET = os.environ["AZURE_CLIENT_SECRET"]
USERNAME      = os.environ["PBI_USERNAME"]
PASSWORD      = os.environ["PBI_PASSWORD"]
WORKSPACE_ID  = os.environ["PBI_WORKSPACE_ID"]

PBI_SCOPE  = "https://analysis.windows.net/powerbi/api/.default"
TOKEN_URL  = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
PBI_BASE   = "https://api.powerbi.com/v1.0/myorg"

OUTPUT_DIR = Path("data/latest")
HIST_DIR   = Path("data/historico")
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
HIST_DIR.mkdir(parents=True, exist_ok=True)

HOY = datetime.date.today().strftime("%Y-%m-%d")

# ── Auth ──────────────────────────────────────────────────────────────────────
def get_token():
    r = requests.post(TOKEN_URL, data={
        "grant_type":    "password",
        "client_id":     CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "username":      USERNAME,
        "password":      PASSWORD,
        "scope":         PBI_SCOPE,
    })
    r.raise_for_status()
    return r.json()["access_token"]

# ── PBI helpers ───────────────────────────────────────────────────────────────
def pbi_get(token, path):
    r = requests.get(f"{PBI_BASE}/{path}", headers={"Authorization": f"Bearer {token}"})
    r.raise_for_status()
    return r.json()

def pbi_dax(token, dataset_id, query):
    r = requests.post(
        f"{PBI_BASE}/groups/{WORKSPACE_ID}/datasets/{dataset_id}/executeQueries",
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json={"queries": [{"query": query}], "serializerSettings": {"includeNulls": True}}
    )
    if not r.ok:
        print(f"DAX error: {r.status_code} {r.text[:300]}")
        return []
    rows = r.json().get("results", [{}])[0].get("tables", [{}])[0].get("rows", [])
    return rows

def safe_val(rows, col, default="—"):
    if not rows:
        return default
    v = rows[0].get(col)
    return v if v is not None else default

# ── Discover datasets ─────────────────────────────────────────────────────────
def find_datasets(token):
    data = pbi_get(token, f"groups/{WORKSPACE_ID}/datasets")
    return {d["name"]: d["id"] for d in data.get("value", [])}

# ── DAX queries for CxC (PAUNO) ───────────────────────────────────────────────
def query_cxc(token, dataset_id):
    """Intenta obtener KPIs de CxC del dataset."""
    # Primero descubrimos tablas disponibles
    tables_r = requests.get(
        f"{PBI_BASE}/groups/{WORKSPACE_ID}/datasets/{dataset_id}/tables",
        headers={"Authorization": f"Bearer {token}"}
    )
    tables = []
    if tables_r.ok:
        tables = [t["name"] for t in tables_r.json().get("value", [])]
    print(f"  Tablas disponibles: {tables}")

    result = {}

    # Query ventas
    ventas_q = """
    EVALUATE
    ROW(
      "ventas_mes_actual", [Ventas Mes Actual],
      "ventas_mes_anterior", [Ventas Mes Anterior]
    )
    """
    rows = pbi_dax(token, dataset_id, ventas_q)
    if rows:
        result["ventas_actual"] = rows[0].get("[ventas_mes_actual]") or rows[0].get("ventas_mes_actual")
        result["ventas_anterior"] = rows[0].get("[ventas_mes_anterior]") or rows[0].get("ventas_mes_anterior")

    # Query mora y CxC
    mora_q = """
    EVALUATE
    ROW(
      "morosidad", [Morosidad],
      "cxc_total", [CxC Total],
      "mas_30_dias", [Mas 30 Dias],
      "rotacion_cxc", [Rotacion CxC]
    )
    """
    rows2 = pbi_dax(token, dataset_id, mora_q)
    if rows2:
        for k,v in rows2[0].items():
            result[k.strip("[]").lower().replace(" ","_")] = v

    return result

# ── Genera summaries.json ─────────────────────────────────────────────────────
def build_summaries(token, datasets):
    summary = {
        "fecha": datetime.date.today().strftime("%d/%m/%Y"),
        "fecha_actualizacion": HOY,
        "generado_por": "JUANITO — Power BI Direct",
        "semaforos": {},
        "semaforo_razon": {},
        "holding_ventas": "—",
        "holding_mora": "—",
        "agenda_ceo": {"decidir_hoy": [], "escalar_semana": []},
        "empresas": {}
    }

    empresa_map = {
        "PAUNO": ["PAUNO (BI)", "PAUNO", "Pauno", "pauno"],
        "AMAUTA": ["AMAUTA (BI)", "AMAUTA", "Amauta"],
        "NEOPACK": ["NEOPACK (BI)", "NEOPACK", "Neopack"],
    }

    for empresa, posibles_nombres in empresa_map.items():
        dataset_id = None
        for nombre in posibles_nombres:
            if nombre in datasets:
                dataset_id = datasets[nombre]
                print(f"  {empresa} → dataset '{nombre}' ({dataset_id})")
                break

        if not dataset_id:
            print(f"  {empresa} → dataset NO encontrado. Disponibles: {list(datasets.keys())}")
            summary["semaforos"][empresa] = "yellow"
            summary["semaforo_razon"][empresa] = "Dataset no conectado aún"
            summary["empresas"][empresa] = {"reportes": {}}
            continue

        kpis = query_cxc(token, dataset_id)
        print(f"  {empresa} KPIs: {kpis}")

        mora_raw = kpis.get("morosidad")
        mora_pct = None
        if mora_raw is not None:
            try:
                mora_pct = float(str(mora_raw).replace("%","").replace(",","."))
            except:
                pass

        sem = "green"
        razon = "Operando dentro de parámetros"
        if mora_pct is not None:
            if mora_pct >= 20:
                sem = "red"
                razon = f"Morosidad crítica {mora_pct:.0f}%"
            elif mora_pct >= 15:
                sem = "yellow"
                razon = f"Morosidad en {mora_pct:.0f}% — sobre meta 15%"

        summary["semaforos"][empresa] = sem
        summary["semaforo_razon"][empresa] = razon

        def fmt_soles(v):
            if v is None: return "—"
            try: return f"S/{float(v):,.0f}"
            except: return str(v)

        summary["empresas"][empresa] = {
            "reportes": {
                "cuentas_por_cobrar": {
                    "estado": sem,
                    "alerta": razon,
                    "kpis": [
                        {"label": "Morosidad", "valor": f"{mora_pct:.0f}%" if mora_pct else str(mora_raw or "—"), "meta": "15%"},
                        {"label": "Rotación CxC", "valor": str(kpis.get("rotacion_cxc") or "—") + " días"},
                        {"label": "CxC Total", "valor": fmt_soles(kpis.get("cxc_total"))},
                        {"label": "Vencido >30d", "valor": fmt_soles(kpis.get("mas_30_dias"))},
                    ]
                },
                "ventas": {
                    "estado": "green",
                    "kpis": [
                        {"label": "Ventas mes actual", "valor": fmt_soles(kpis.get("ventas_actual"))},
                        {"label": "Ventas mes anterior", "valor": fmt_soles(kpis.get("ventas_anterior"))},
                    ]
                }
            }
        }

        # Holding totals (suma)
        if mora_pct:
            summary["holding_mora"] = f"{mora_pct:.0f}% (PAUNO)"

        # Agenda CEO automática
        if sem == "red":
            summary["agenda_ceo"]["decidir_hoy"].append({
                "empresa": empresa,
                "texto": f"Intervenir cobranza {empresa}: mora {mora_pct:.0f}%, vencido >30d {fmt_soles(kpis.get('mas_30_dias'))}",
                "responsable": "Gerencia Financiera"
            })

    return summary

# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    print("=== JUANITO Power BI Fetcher ===")
    print(f"Fecha: {HOY}")

    print("\n[1] Autenticando en Power BI...")
    try:
        token = get_token()
        print("    Token OK")
    except Exception as e:
        print(f"    ERROR auth: {e}")
        sys.exit(1)

    print("\n[2] Descubriendo datasets...")
    datasets = find_datasets(token)
    print(f"    Datasets encontrados: {list(datasets.keys())}")

    print("\n[3] Generando summaries.json...")
    summaries = build_summaries(token, datasets)
    out_path = OUTPUT_DIR / "summaries.json"
    out_path.write_text(json.dumps(summaries, ensure_ascii=False, indent=2))
    print(f"    Guardado: {out_path}")

    # Actualizar historico
    hist_path = HIST_DIR / "historico.json"
    try:
        historico = json.loads(hist_path.read_text()) if hist_path.exists() else {
            "generado_por": "JUANITO — Módulo Histórico",
            "ultima_actualizacion": HOY,
            "meses": [],
            "empresas": {"PAUNO": {"meses": []}, "AMAUTA": {"meses": []}, "NEOPACK": {"meses": []}}
        }
        mes_label = datetime.date.today().strftime("%b-%y")
        for emp, ed in summaries.get("empresas", {}).items():
            meses = historico["empresas"].setdefault(emp, {"meses": []})["meses"]
            ventas_kpi = next((k["valor"] for k in (ed.get("reportes",{}).get("ventas",{}).get("kpis",[]) or []) if "actual" in k.get("label","").lower()), "—")
            mora_kpi = next((k["valor"] for k in (ed.get("reportes",{}).get("cuentas_por_cobrar",{}).get("kpis",[]) or []) if "mora" in k.get("label","").lower()), "—")
            entrada = {"mes": mes_label, "ventas": ventas_kpi, "mora": mora_kpi}
            if not meses or meses[-1]["mes"] != mes_label:
                meses.append(entrada)
                if len(meses) > 6:
                    meses.pop(0)
        historico["ultima_actualizacion"] = HOY
        hist_path.write_text(json.dumps(historico, ensure_ascii=False, indent=2))
        print(f"    Histórico actualizado: {hist_path}")
    except Exception as e:
        print(f"    Histórico: {e}")

    print("\n=== Completado ===")

if __name__ == "__main__":
    main()
