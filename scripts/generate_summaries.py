#!/usr/bin/env python3
"""
generate_summaries.py -- Juanito Garcia
Lee los PNGs del scraper, llama a Claude API y genera summaries.json
"""

import os
import sys
import json
import base64
import requests
from pathlib import Path
from datetime import datetime

ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
DATA_DIR = Path("data/latest")
API_URL = "https://api.anthropic.com/v1/messages"

REPORTES = {
    "PAUNO": [
        ("01_cuentas_por_cobrar.png",  "cuentas_por_cobrar"),
        ("02_cuentas_por_pagar.png",   "cuentas_por_pagar"),
        ("03_margen_variable.png",     "margen_variable"),
        ("04_mermas.png",              "mermas"),
        ("05_compras.png",             "compras"),
        ("07_control_interno.png",     "control_interno"),
        ("10_sop_inventario.png",      "sop_inventario"),
        ("11_fill_rate.png",           "fill_rate"),
        ("12_productividad.png",       "productividad"),
        ("13_margen_variable_pag2.png","margen_variable_pag2"),
    ],
    "AMAUTA": [
        ("01_reporte_01.png", "cuentas_por_cobrar"),
        ("02_reporte_02.png", "cuentas_por_pagar"),
        ("03_reporte_03.png", "inventario"),
    ],
    "NEOPACK": [
        ("01_reporte_01.png", "cuentas_por_cobrar"),
        ("02_reporte_02.png", "cuentas_por_pagar"),
        ("03_reporte_03.png", "ventas"),
    ],
}

SEMAFORO_PROMPT = (
    "Analiza TODOS estos reportes del holding (PAUNO, AMAUTA, NEOPACK) y responde "
    "SOLO con JSON valido, sin texto adicional.\n\n"
    "REGLAS:\n"
    "- Lee numeros directamente de las imagenes. Si no ves claro escribe null.\n"
    "- Compras ratio = Consumo/Compra: >100% jala stock, 80-100% normal, <80% sobre-compra.\n"
    "- Semaforo: red=KPI critico fuera de meta, yellow=alerta, green=ok\n\n"
    'Responde con este JSON exacto:\n'
    '{\n'
    '  "fecha": "DD/MM/YYYY",\n'
    '  "hora": "HH:MM",\n'
    '  "semaforos": {"PAUNO": "red|yellow|green", "AMAUTA": "red|yellow|green", "NEOPACK": "red|yellow|green"},\n'
    '  "semaforo_razon": {"PAUNO": "KPI mas critico con numero", "AMAUTA": "...", "NEOPACK": "..."},\n'
    '  "holding_ventas": "S/X.XM",\n'
    '  "holding_mora": "XX%",\n'
    '  "agenda_ceo": {\n'
    '    "decidir_hoy": [{"texto": "accion concreta", "empresa": "PAUNO|AMAUTA|NEOPACK", "responsable": "area"}],\n'
    '    "escalar_semana": [{"texto": "seguimiento necesario", "empresa": "...", "responsable": "..."}],\n'
    '    "monitorear": [{"texto": "tendencia a vigilar", "empresa": "..."}]\n'
    '  }\n'
    '}'
)

REPORTE_PROMPT = {
    "cuentas_por_cobrar": (
        "Analiza esta imagen de Cuentas por Cobrar y responde SOLO con JSON valido:\n"
        '{\n'
        '  "alerta": "frase de alerta si hay problema, null si todo ok",\n'
        '  "estado": "red|yellow|green",\n'
        '  "kpis": [\n'
        '    {"label": "Morosidad", "valor": "XX%", "meta": "<15%", "estado": "red|yellow|green", "tendencia": "sube|baja|estable"},\n'
        '    {"label": "CxC total", "valor": "S/X.XM", "meta": null, "estado": "green", "tendencia": "estable"},\n'
        '    {"label": "Rotacion", "valor": "XXd", "meta": null, "estado": "green", "tendencia": "estable"}\n'
        '  ],\n'
        '  "tramos": [\n'
        '    {"label": "Por vencer", "valor": "S/X.XM", "pct": "XX%", "estado": "green"},\n'
        '    {"label": "0-30 dias", "valor": "S/XXXK", "pct": "XX%", "estado": "yellow"},\n'
        '    {"label": "+30 dias", "valor": "S/XXXK", "pct": "XX%", "estado": "red"}\n'
        '  ],\n'
        '  "riesgos": [\n'
        '    {"nombre": "nombre cliente", "canal": "canal", "total": "S/XXXK", "vencido30": "S/XXXK", "estado": "red|yellow"}\n'
        '  ],\n'
        '  "tendencia_mora": [{"mes": "Ene", "pct": 11}],\n'
        '  "acciones": ["accion 1", "accion 2", "accion 3"]\n'
        '}'
    ),
    "cuentas_por_pagar": (
        "Analiza esta imagen de Cuentas por Pagar y responde SOLO con JSON valido:\n"
        '{\n'
        '  "alerta": "frase de alerta si hay problema, null si ok",\n'
        '  "estado": "red|yellow|green",\n'
        '  "kpis": [\n'
        '    {"label": "Dias CxP", "valor": "XXXd", "meta": "<90d", "estado": "red|yellow|green", "tendencia": "sube|baja|estable"},\n'
        '    {"label": "CxP total", "valor": "S/XX.XM", "meta": null, "estado": "yellow", "tendencia": "sube"},\n'
        '    {"label": "Refinanciado", "valor": "S/X.XM", "meta": null, "estado": "yellow", "tendencia": "estable"}\n'
        '  ],\n'
        '  "tramos": [\n'
        '    {"label": "Vigente", "valor": "S/X.XM", "pct": "XX%", "estado": "green"},\n'
        '    {"label": "Venc. <15d", "valor": "S/X.XM", "pct": "XX%", "estado": "yellow"},\n'
        '    {"label": "Venc. 16-90d", "valor": "S/X.XM", "pct": "XX%", "estado": "red"}\n'
        '  ],\n'
        '  "proveedores_criticos": [\n'
        '    {"nombre": "nombre proveedor", "categoria": "categoria", "monto": "S/X.XM", "estado": "red|yellow"}\n'
        '  ],\n'
        '  "tendencia_dias": [{"mes": "Ago", "dias": 42}],\n'
        '  "acciones": ["accion 1", "accion 2", "accion 3"]\n'
        '}'
    ),
    "margen_variable": (
        "Analiza esta imagen de Margen Variable y responde SOLO con JSON valido:\n"
        '{\n'
        '  "alerta": "frase si margen < 46%, null si ok",\n'
        '  "estado": "red|yellow|green",\n'
        '  "kpis": [\n'
        '    {"label": "Margen total", "valor": "XX%", "meta": "52%", "estado": "red|yellow|green", "tendencia": "sube|baja|estable"},\n'
        '    {"label": "Ventas mes", "valor": "S/X.XM", "meta": null, "estado": "green", "tendencia": "estable"},\n'
        '    {"label": "Precio/kg", "valor": "S/X.XX", "meta": null, "estado": "green", "tendencia": "estable"}\n'
        '  ],\n'
        '  "por_uen": [\n'
        '    {"uen": "B&D", "margen": "XX%", "estado": "red|yellow|green"},\n'
        '    {"uen": "MAQUILA", "margen": "XX%", "estado": "red|yellow|green"},\n'
        '    {"uen": "TIGO", "margen": "XX%", "estado": "red|yellow|green"}\n'
        '  ],\n'
        '  "tendencia_margen": [{"mes": "Ene", "pct": 52.85}],\n'
        '  "acciones": ["accion 1", "accion 2"]\n'
        '}'
    ),
    "mermas": (
        "Analiza esta imagen de Mermas y responde SOLO con JSON valido:\n"
        '{\n'
        '  "alerta": "frase si alguna planta > 3%, null si ok",\n'
        '  "estado": "red|yellow|green",\n'
        '  "kpis": [\n'
        '    {"label": "Merma Ate", "valor": "X.X%", "meta": "2%", "estado": "red|yellow|green", "tendencia": "sube|baja|estable"},\n'
        '    {"label": "Merma Pachacamac", "valor": "X.X%", "meta": "2%", "estado": "red|yellow|green", "tendencia": "sube|baja|estable"},\n'
        '    {"label": "Merma Terceros", "valor": "X.X%", "meta": "2%", "estado": "red|yellow|green", "tendencia": "sube|baja|estable"}\n'
        '  ],\n'
        '  "tendencia_ate": [{"mes": "Ene", "pct": 2.21}],\n'
        '  "tendencia_pachacamac": [{"mes": "Ene", "pct": 1.59}],\n'
        '  "acciones": ["accion 1", "accion 2"]\n'
        '}'
    ),
    "compras": (
        "Analiza esta imagen de Compras (Ratio Consumo/Compra) y responde SOLO con JSON valido.\n"
        "REGLA: ratio 80-100% = normal (green), >100% = jala stock, <80% = sobre-compra.\n"
        '{\n'
        '  "alerta": "frase si ratio extremo, null si normal",\n'
        '  "estado": "red|yellow|green",\n'
        '  "kpis": [\n'
        '    {"label": "Ratio agosto", "valor": "XXX%", "meta": "80-100%", "estado": "red|yellow|green", "tendencia": "sube|baja|estable"},\n'
        '    {"label": "Ratio julio", "valor": "XX%", "meta": "80-100%", "estado": "green", "tendencia": "estable"},\n'
        '    {"label": "Ratio promedio", "valor": "XX%", "meta": "80-100%", "estado": "green", "tendencia": "estable"}\n'
        '  ],\n'
        '  "por_categoria": [\n'
        '    {"categoria": "Materia Prima", "ratio": "XX%", "interpretacion": "normal|jala stock|sobre-compra", "estado": "red|yellow|green"},\n'
        '    {"categoria": "Envases y Embalajes", "ratio": "XX%", "interpretacion": "...", "estado": "..."},\n'
        '    {"categoria": "Suministros", "ratio": "XX%", "interpretacion": "...", "estado": "..."},\n'
        '    {"categoria": "Repuestos", "ratio": "XX%", "interpretacion": "...", "estado": "..."}\n'
        '  ],\n'
        '  "tendencia_ratio": [{"mes": "Ago", "ratio": 92.45}],\n'
        '  "acciones": ["accion 1", "accion 2"]\n'
        '}'
    ),
    "control_interno": (
        "Analiza esta imagen de Control Interno y responde SOLO con JSON valido:\n"
        '{\n'
        '  "alerta": "frase si cumplimiento < 70%, null si ok",\n'
        '  "estado": "red|yellow|green",\n'
        '  "kpis": [\n'
        '    {"label": "Cumplimiento", "valor": "XX%", "meta": "85%", "estado": "red|yellow|green", "tendencia": "sube|baja|estable"},\n'
        '    {"label": "Puntos criticos", "valor": "XX", "meta": "0", "estado": "red", "tendencia": "estable"},\n'
        '    {"label": "Satisfactorio", "valor": "XX", "meta": null, "estado": "green", "tendencia": "estable"}\n'
        '  ],\n'
        '  "por_gerente": [{"nombre": "nombre", "pct": "XX%", "estado": "red|yellow|green"}],\n'
        '  "tendencia": [{"mes": "Ene", "pct": 44.44}],\n'
        '  "acciones": ["accion 1", "accion 2"]\n'
        '}'
    ),
    "sop_inventario": (
        "Analiza esta imagen de S&OP Inventario y responde SOLO con JSON valido:\n"
        '{\n'
        '  "alerta": "frase si dead stock > S/1M, null si ok",\n'
        '  "estado": "red|yellow|green",\n'
        '  "kpis": [\n'
        '    {"label": "Dead stock", "valor": "S/X.XM", "meta": "<S/500K", "estado": "red|yellow|green", "tendencia": "sube|baja|estable"},\n'
        '    {"label": "% Dead", "valor": "XX%", "meta": "<5%", "estado": "red|yellow|green", "tendencia": "estable"},\n'
        '    {"label": "Working", "valor": "XX%", "meta": ">70%", "estado": "green", "tendencia": "estable"}\n'
        '  ],\n'
        '  "clasificacion": [\n'
        '    {"tipo": "Working", "monto": "S/X.XM", "pct": "XX%", "estado": "green"},\n'
        '    {"tipo": "Exceso 1", "monto": "S/XXXK", "pct": "XX%", "estado": "yellow"},\n'
        '    {"tipo": "Dead", "monto": "S/X.XM", "pct": "XX%", "estado": "red"}\n'
        '  ],\n'
        '  "dead_por_categoria": [\n'
        '    {"categoria": "Envases y Embalajes", "monto": "S/XXXK", "pct_dead": "XX%", "estado": "red"}\n'
        '  ],\n'
        '  "acciones": ["accion 1", "accion 2"]\n'
        '}'
    ),
    "fill_rate": (
        "Analiza esta imagen de Fill Rate y responde SOLO con JSON valido:\n"
        '{\n'
        '  "alerta": "frase si fill rate < 85%, null si ok",\n'
        '  "estado": "red|yellow|green",\n'
        '  "kpis": [\n'
        '    {"label": "Fill Rate", "valor": "XX%", "meta": "98%", "estado": "red|yellow|green", "tendencia": "sube|baja|estable"},\n'
        '    {"label": "Brecha", "valor": "-Xpp", "meta": "0", "estado": "red|yellow|green", "tendencia": "estable"},\n'
        '    {"label": "Clientes afect.", "valor": "XX", "meta": null, "estado": "yellow", "tendencia": "estable"}\n'
        '  ],\n'
        '  "acciones": ["accion 1", "accion 2"]\n'
        '}'
    ),
    "productividad": (
        "Analiza esta imagen de Productividad y responde SOLO con JSON valido:\n"
        '{\n'
        '  "alerta": null,\n'
        '  "estado": "red|yellow|green",\n'
        '  "kpis": [\n'
        '    {"label": "S/kg producido", "valor": "S/X.XX", "meta": null, "estado": "green", "tendencia": "sube|baja|estable"},\n'
        '    {"label": "vs 2025", "valor": "S/X.XX", "meta": null, "estado": "green", "tendencia": "baja"},\n'
        '    {"label": "Planilla total", "valor": "S/X.XM", "meta": null, "estado": "green", "tendencia": "estable"}\n'
        '  ],\n'
        '  "tendencia": [{"mes": "Ene", "val": 0.51}],\n'
        '  "acciones": ["accion 1"]\n'
        '}'
    ),
    "margen_variable_pag2": (
        "Analiza esta imagen de Avance de Ventas vs Presupuesto y responde SOLO con JSON valido:\n"
        '{\n'
        '  "alerta": "frase si avance < 70% a mitad de mes, null si ok",\n'
        '  "estado": "red|yellow|green",\n'
        '  "kpis": [\n'
        '    {"label": "Avance total", "valor": "XX%", "meta": "proporcional al dia", "estado": "red|yellow|green", "tendencia": "sube|baja|estable"},\n'
        '    {"label": "Ventas reales", "valor": "S/XXXK", "meta": null, "estado": "green", "tendencia": "estable"},\n'
        '    {"label": "Dia del mes", "valor": "D/31", "meta": null, "estado": "green", "tendencia": "estable"}\n'
        '  ],\n'
        '  "por_canal": [\n'
        '    {"canal": "Moderno", "avance": "XX%", "real": "S/XXXK", "estado": "red|yellow|green"},\n'
        '    {"canal": "Tradicional", "avance": "XX%", "real": "S/XXXK", "estado": "red|yellow|green"},\n'
        '    {"canal": "Horeca", "avance": "XX%", "real": "S/XXXK", "estado": "red|yellow|green"}\n'
        '  ],\n'
        '  "por_marca": [\n'
        '    {"marca": "MAQUILA", "avance": "XX%", "real": "S/XXXK", "estado": "green"},\n'
        '    {"marca": "B&D", "avance": "XX%", "real": "S/XXXK", "estado": "yellow"},\n'
        '    {"marca": "TIGO", "avance": "XX%", "real": "S/XXXK", "estado": "red"}\n'
        '  ],\n'
        '  "acciones": ["accion 1", "accion 2"]\n'
        '}'
    ),
    "inventario": (
        "Analiza esta imagen de Inventario y responde SOLO con JSON valido:\n"
        '{\n'
        '  "alerta": "frase si dead > 5% del total, null si ok",\n'
        '  "estado": "red|yellow|green",\n'
        '  "kpis": [\n'
        '    {"label": "Stock total", "valor": "S/X.XM", "meta": null, "estado": "green", "tendencia": "estable"},\n'
        '    {"label": "Cobertura", "valor": "XXd", "meta": null, "estado": "green", "tendencia": "estable"},\n'
        '    {"label": "Dead stock", "valor": "X%", "meta": "<5%", "estado": "red|yellow|green", "tendencia": "estable"}\n'
        '  ],\n'
        '  "acciones": ["accion 1"]\n'
        '}'
    ),
    "ventas": (
        "Analiza esta imagen de Ventas y responde SOLO con JSON valido:\n"
        '{\n'
        '  "alerta": "frase si avance < 60% a mitad de mes, null si ok",\n'
        '  "estado": "red|yellow|green",\n'
        '  "kpis": [\n'
        '    {"label": "Avance", "valor": "XX%", "meta": "100%", "estado": "red|yellow|green", "tendencia": "sube|baja|estable"},\n'
        '    {"label": "Facturacion", "valor": "S/XXXK", "meta": null, "estado": "green", "tendencia": "estable"},\n'
        '    {"label": "Categoria lider", "valor": "categoria", "meta": null, "estado": "green", "tendencia": "estable"}\n'
        '  ],\n'
        '  "acciones": ["accion 1"]\n'
        '}'
    ),
}


def img_b64(path: Path) -> str:
    data = base64.b64encode(path.read_bytes()).decode("ascii")
    return data


def llamar_api(content: list, max_tokens: int = 1500) -> dict:
    headers = {
        "x-api-key": ANTHROPIC_API_KEY,
        "anthropic-version": "2023-06-01",
        "content-type": "application/json",
    }
    body = json.dumps(
        {"model": "claude-sonnet-4-5", "max_tokens": max_tokens,
         "messages": [{"role": "user", "content": content}]},
        ensure_ascii=True,
    ).encode("ascii")

    for intento in range(3):
        try:
            resp = requests.post(API_URL, headers=headers, data=body, timeout=120)
            resp.raise_for_status()
            texto = resp.json()["content"][0]["text"].strip()
            if texto.startswith("```"):
                texto = texto.split("```")[1]
                if texto.startswith("json"):
                    texto = texto[4:]
            return json.loads(texto)
        except Exception as e:
            print(f"  Error intento {intento + 1}: {e}")
            if intento == 2:
                return {}
    return {}


def generar_semaforo_y_agenda(data_dir: Path) -> dict:
    print("Generando semaforo global y agenda CEO...")
    content = []
    for empresa, reportes in REPORTES.items():
        for archivo, _ in reportes[:3]:
            img_path = data_dir / empresa / archivo
            if img_path.exists():
                content.append({"type": "text", "text": "--- " + empresa + " / " + archivo + " ---"})
                content.append({"type": "image", "source": {"type": "base64", "media_type": "image/png", "data": img_b64(img_path)}})
    content.append({"type": "text", "text": SEMAFORO_PROMPT})
    return llamar_api(content, max_tokens=2000)


def generar_reporte(empresa: str, archivo: str, tipo: str, data_dir: Path) -> dict:
    img_path = data_dir / empresa / archivo
    if not img_path.exists():
        print(f"  NO ENCONTRADO: {img_path}")
        return {}
    print(f"  Analizando {empresa}/{archivo}...")
    prompt = REPORTE_PROMPT.get(tipo, REPORTE_PROMPT["cuentas_por_cobrar"])
    content = [
        {"type": "text", "text": "Empresa: " + empresa + " | Reporte: " + tipo},
        {"type": "image", "source": {"type": "base64", "media_type": "image/png", "data": img_b64(img_path)}},
        {"type": "text", "text": prompt},
    ]
    return llamar_api(content, max_tokens=1200)


def main():
    if not ANTHROPIC_API_KEY:
        print("ERROR: ANTHROPIC_API_KEY no configurada")
        sys.exit(1)

    global_data = generar_semaforo_y_agenda(DATA_DIR)

    empresas_data = {}
    for empresa, reportes in REPORTES.items():
        print(f"\nAnalizando {empresa}...")
        empresas_data[empresa] = {"reportes": {}}
        for archivo, tipo in reportes:
            reporte_data = generar_reporte(empresa, archivo, tipo, DATA_DIR)
            empresas_data[empresa]["reportes"][tipo] = reporte_data

    resultado = {
        **global_data,
        "generado_en": datetime.now().isoformat(),
        "empresas": empresas_data,
    }

    out = DATA_DIR / "summaries.json"
    out.write_bytes(json.dumps(resultado, ensure_ascii=False, indent=2).encode("utf-8"))
    print(f"\nSummaries guardados en {out}")


if __name__ == "__main__":
    main()
