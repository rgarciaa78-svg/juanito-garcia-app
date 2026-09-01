#!/usr/bin/env python3
"""
Prueba nombres de medidas exactos extraídos de la captura del reporte CxP.
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

# CxP dataset
CXP_ID   = "45a8ab8d-e162-4398-a260-3a9a5f90829f"
PROD_ID  = "a042ba6a-c82c-4fc1-bf95-b9e84bd15fc6"
PLAN_ID  = "30074d92-7ec1-4762-82f2-1cb29c15dcfe"
CONS_ID  = "c972c8cb-e5fc-4b60-8f5e-265a78e1e796"

# Candidatos CxP deducidos de la captura
CXP_CANDIDATES = [
    # Total
    "Cuentas x Pagar", "CUENTAS X PAGAR", "Cuentas por Pagar", "Total CxP",
    "CxP Total", "Total Cuentas Pagar", "Saldo CxP", "Total Saldo",
    # Vigente
    "Vigente", "CxP Vigente", "Saldo Vigente", "No Vencido",
    # Vencido tramos
    "Vencido", "Total Vencido", "Saldo Vencido",
    "Vencido 0-15", "Vencido <15", "Vencido Menor 15",
    "Vencido 16-30", "Vencido 30-60", "Vencido 31-90",
    "Vencido 31 a 90", "Vencido 16 a 30",
    # Refinanciado
    "Refinanciado", "Con Cronograma", "Refinanciamiento",
    # Dias
    "Dias CxP", "DPP", "Dias de Pago", "Días de Pago",
    "Rotación CxP", "Rotacion CxP",
    # Número proveedores
    "# Proveedores", "Proveedores", "Nro Proveedores",
    # Compras (eje del gráfico)
    "Compras CxP", "Total Compras CxP",
]

print("=== CxP — probando candidatos ===")
found_cxp = {}
for name in CXP_CANDIDATES:
    v, ok = try_m(CXP_ID, name)
    if ok:
        found_cxp[name] = v
        print(f"  ✓ [{name}] = {v}")
    else:
        print(f"  ✗ [{name}]")

print(f"\nEncontradas CxP: {found_cxp}")

# Candidatos Productividad deducidos del nombre del reporte
# "Reporte de productividad por funcionario" + página "R. PRODUCTIVIDAD VAL"
PROD_CANDIDATES = [
    # Planilla / costo MOD
    "Planilla", "Total Planilla", "Costo Planilla", "Gasto Planilla",
    "MOD", "Costo MOD", "Total MOD",
    "Planilla kg", "Planilla/kg", "S/.kg", "S/.Kg", "S/kg",
    "Costo x kg", "Costo por kg", "S/ x kg",
    # Kg producidos
    "Kg", "Total Kg", "Kg Producidos", "Kg Produccion", "Kg Producción",
    "Produccion", "Producción", "Total Produccion",
    # Productividad
    "Productividad", "Indice Productividad", "Ratio Productividad",
    "Eficiencia", "% Eficiencia",
    # Por funcionario
    "Funcionario", "Total Funcionarios",
    # Costo valorizado
    "Planilla Val", "MOD Valorizado", "Costo Valorizado",
    "VAL", "Valorizado",
]

print("\n=== Productividad — probando candidatos ===")
found_prod = {}
for name in PROD_CANDIDATES:
    v, ok = try_m(PROD_ID, name)
    if ok:
        found_prod[name] = v
        print(f"  ✓ [{name}] = {v}")
    else:
        print(f"  ✗ [{name}]")

print(f"\nEncontradas Prod: {found_prod}")

# Fill Rate en planificacion — página "REPORTE DE S&OP"
FR_CANDIDATES = [
    "Fill Rate", "% Fill Rate", "FR", "% FR",
    "Nivel Servicio", "% Nivel Servicio", "Nivel de Servicio",
    "Atendido", "% Atendido", "Pedidos Atendidos",
    "S&OP", "REPORTE S&OP",
    "Cumplimiento Pedidos", "% Cumplimiento Pedidos",
    "OC Atendidas", "% OC Atendidas",
    "Demanda Atendida", "% Demanda",
    # Consumo ratio
    "Ratio Consumo", "Ratio C/C", "Consumo/Compra", "Ratio",
]

print("\n=== Planificacion — Fill Rate y S&OP ===")
found_plan = {}
for name in FR_CANDIDATES:
    v, ok = try_m(PLAN_ID, name)
    if ok:
        found_plan[name] = v
        print(f"  ✓ [{name}] = {v}")
    else:
        print(f"  ✗ [{name}]")

print(f"\nEncontradas Plan: {found_plan}")

# Consumo
CONS_CANDIDATES = [
    "Consumo", "Total Consumo", "Importe Consumo",
    "Compras", "Total Compras",
    "Ratio", "Ratio Consumo",
]
print("\n=== Consumo ===")
found_cons = {}
for name in CONS_CANDIDATES:
    v, ok = try_m(CONS_ID, name)
    if ok:
        found_cons[name] = v
        print(f"  ✓ [{name}] = {v}")
    else:
        print(f"  ✗ [{name}]")

print(f"\nEncontradas Consumo: {found_cons}")
print("\n=== FIN ===")
