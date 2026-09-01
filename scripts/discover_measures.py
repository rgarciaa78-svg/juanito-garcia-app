#!/usr/bin/env python3
"""Prueba nombres exactos de Productividad y Fill Rate desde capturas del reporte."""
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
H = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

def try_m(ds_id, name):
    url = f"{PBI_BASE}/groups/{WS_ID}/datasets/{ds_id}/executeQueries"
    resp = requests.post(url, headers=H,
                         json={"queries": [{"query": f'EVALUATE ROW("v", [{name}])'}],
                               "serializerSettings": {"includeNulls": True}},
                         timeout=15)
    if resp.ok:
        rows = resp.json().get("results", [{}])[0].get("tables", [{}])[0].get("rows", [])
        v = list(rows[0].values())[0] if rows else None
        return v, True
    return None, False

PROD_ID = "a042ba6a-c82c-4fc1-bf95-b9e84bd15fc6"
PLAN_ID = "30074d92-7ec1-4762-82f2-1cb29c15dcfe"

# Candidatos Productividad exactos desde la captura
PROD_CANDIDATES = [
    # Producción KG
    "Produccion Total (KG)", "PRODUCCION TOTAL (KG)", "Produccion Total",
    "Producción Total (KG)", "Producción Total",
    "KG Total", "Total KG", "KG Produccion", "KG Producción",
    # Venta Neta KG
    "Venta Neta (KG)", "VENTA NETA (KG)", "Venta Neta",
    "Ventas Neta KG", "Venta KG", "KG Vendidos",
    # Planilla
    "Planilla Total (S/.)", "PLANILLA TOTAL (S/.)", "Planilla Total",
    "Planilla (S/.)", "Total Planilla S/.", "Planilla S/.",
    # Ratio S/kg
    "Planilla (S/.) entre Kg Producido", "PLANILLA (S/.) ENTRE KG PRODUCIDO",
    "Planilla entre KG", "S/. x KG", "S/.xKG", "S/Kg", "Planilla/KG",
    "Ratio Planilla KG", "Costo Planilla KG",
]

print("=== PRODUCTIVIDAD ===")
found_prod = {}
for name in PROD_CANDIDATES:
    v, ok = try_m(PROD_ID, name)
    if ok:
        found_prod[name] = v
        print(f"  ✓ [{name}] = {v}")
    else:
        print(f"  ✗ [{name}]")

print(f"\nEncontradas: {found_prod}")

# Fill Rate — en el reporte de Planificaciones (11) hay página "REPORTE DE S&OP"
# y "ANALISIS DE COMPRA" — probar nombres que encajan con S&OP
FR_CANDIDATES = [
    # Fill Rate variantes
    "Fill Rate", "% Fill Rate", "Fill Rate (%)", "FR", "% FR",
    "Fill Rate Unidades", "Fill Rate Valor",
    # S&OP
    "Nivel Servicio", "Nivel de Servicio", "% Nivel Servicio",
    "Atendido %", "% Atendido", "Demanda Atendida", "% Demanda Atendida",
    "Pedidos Atendidos", "% Pedidos Atendidos", "Cumplimiento Pedidos",
    # Análisis compra — posibles medidas
    "Ratio C/C", "Ratio Consumo", "Ratio Compras",
    "Consumo/Compra", "C/C",
]

print("\n=== FILL RATE (Planificacion) ===")
found_fr = {}
for name in FR_CANDIDATES:
    v, ok = try_m(PLAN_ID, name)
    if ok:
        found_fr[name] = v
        print(f"  ✓ [{name}] = {v}")
    else:
        print(f"  ✗ [{name}]")

print(f"\nEncontradas FR: {found_fr}")
print("\n=== FIN ===")
