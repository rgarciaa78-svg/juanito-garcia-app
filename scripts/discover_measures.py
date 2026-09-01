#!/usr/bin/env python3
"""
Descubre medidas via páginas/visuales de reportes y Export API.
"""
import os, json, requests, time

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
HG = {"Authorization": f"Bearer {token}"}

DATASETS = {
    "cxp":           "45a8ab8d-e162-4398-a260-3a9a5f90829f",
    "productividad": "a042ba6a-c82c-4fc1-bf95-b9e84bd15fc6",
    "planificacion": "30074d92-7ec1-4762-82f2-1cb29c15dcfe",
    "consumo":       "c972c8cb-e5fc-4b60-8f5e-265a78e1e796",
}

# 1. Listar reportes y mapear a dataset
print("\n=== REPORTES ===")
reps = requests.get(f"{PBI_BASE}/groups/{WS_ID}/reports", headers=HG).json().get("value", [])
rep_by_ds = {}
for rep in reps:
    ds = rep.get("datasetId", "")
    ds_key = next((k for k, v in DATASETS.items() if v == ds), None)
    print(f"  {rep['name']} | report_id={rep['id']} | ds_key={ds_key}")
    if ds_key:
        rep_by_ds[ds_key] = rep["id"]

# 2. Para cada reporte: explorar páginas y visuales
print("\n=== PÁGINAS Y VISUALES ===")
for ds_key, rep_id in rep_by_ds.items():
    print(f"\n--- {ds_key} (report={rep_id[:8]}...) ---")
    pages_r = requests.get(f"{PBI_BASE}/groups/{WS_ID}/reports/{rep_id}/pages", headers=HG)
    if not pages_r.ok:
        print(f"  ERROR páginas: {pages_r.status_code}")
        continue
    pages = pages_r.json().get("value", [])
    for page in pages[:3]:  # primeras 3 páginas
        pname = page["name"]
        pdisplay = page.get("displayName", pname)
        print(f"  Página: {pdisplay} ({pname})")
        vis_r = requests.get(
            f"{PBI_BASE}/groups/{WS_ID}/reports/{rep_id}/pages/{pname}/visuals",
            headers=HG)
        if vis_r.ok:
            visuals = vis_r.json().get("value", [])
            for vis in visuals[:10]:
                vtype = vis.get("type", "?")
                vtitle = vis.get("title", "")
                print(f"    Visual: [{vtype}] {vtitle}")
                # Imprimir todo el contenido del visual para ver campos
                if "dataViews" in vis or "fields" in vis or "columns" in vis:
                    print(f"      data: {json.dumps(vis, ensure_ascii=False)[:300]}")
        else:
            print(f"    ERROR visuals: {vis_r.status_code} {vis_r.text[:100]}")

# 3. Intentar Export To File para obtener datos de visuals
print("\n=== EXPORT TO FILE (CSV) ===")
for ds_key, rep_id in rep_by_ds.items():
    print(f"\n--- {ds_key} ---")
    # Exportar primera página a CSV de underlying data
    export_body = {
        "format": "CSV",
        "paginatedReportConfiguration": None,
        "powerBIReportConfiguration": {
            "pages": [{"pageName": None}],  # primera página
        }
    }
    # Intentar exportar a XLSX para ver datos
    export_r = requests.post(
        f"{PBI_BASE}/groups/{WS_ID}/reports/{rep_id}/ExportTo",
        headers=H,
        json={"format": "XLSX"}
    )
    print(f"  ExportTo XLSX: {export_r.status_code} {export_r.text[:200]}")
    if export_r.ok:
        export_id = export_r.json().get("id")
        print(f"  Export ID: {export_id}")
        # Esperar que termine
        for _ in range(10):
            time.sleep(3)
            status_r = requests.get(
                f"{PBI_BASE}/groups/{WS_ID}/reports/{rep_id}/exports/{export_id}",
                headers=HG)
            status = status_r.json().get("status", "?")
            print(f"  Status: {status}")
            if status == "Succeeded":
                # Descargar
                dl_r = requests.get(
                    f"{PBI_BASE}/groups/{WS_ID}/reports/{rep_id}/exports/{export_id}/file",
                    headers=HG)
                print(f"  Archivo: {dl_r.status_code}, size={len(dl_r.content)} bytes")
                break
            elif status in ["Failed", "Undefined"]:
                break

# 4. Intentar Analyze in Excel (genera un archivo ODC con schema)
print("\n=== ANALYZE IN EXCEL (schema) ===")
for ds_key, ds_id in DATASETS.items():
    analyze_r = requests.get(
        f"{PBI_BASE}/groups/{WS_ID}/datasets/{ds_id}/analyzeInExcel",
        headers=HG)
    print(f"  {ds_key}: {analyze_r.status_code} {analyze_r.text[:300]}")

print("\n=== FIN ===")
