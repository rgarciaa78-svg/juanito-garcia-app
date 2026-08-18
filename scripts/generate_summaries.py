#!/usr/bin/env python3
"""
generate_summaries.py -- Juanito Garcia
Lee los PNGs del scraper, llama a Claude API y genera summaries.json
para la web app. Se ejecuta desde GitHub Actions despues del scraper.
"""

import os
import sys
import json
import base64
from pathlib import Path
from datetime import datetime

try:
    import anthropic
except ImportError:
    print("ERROR: pip install anthropic")
    sys.exit(1)

ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
DATA_DIR = Path("data/latest")

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

SEMAFORO_PROMPT = """Analiza TODOS estos reportes del holding (PAUNO, AMAUTA, NEOPACK) y responde SOLO con JSON valido, sin texto adicional.

REGLAS:
- Lee numeros directamente de las imagenes. Si no ves claro escribe null.
- Compras ratio = Consumo/Compra: >100% jala stock, 80-100% normal, <80% sobre-compra. NUNCA marques >95% como malo.
- Semaforo: "red"=KPI critico fuera de meta, "yellow"=alerta o tendencia negativa, "green"=ok o mejorando

Responde con este JSON exacto:
{
  "fecha": "DD/MM/YYYY",
  "hora": "HH:MM",
  "semaforos": {"PAUNO": "red|yellow|green", "AMAUTA": "red|yellow|green", "NEOPACK": "red|yellow|green"},
  "semaforo_razon": {"PAUNO": "KPI mas critico con numero", "AMAUTA": "...", "NEOPACK": "..."},
  "holding_ventas": "S/X.XM",
  "holding_mora": "XX%",
  "agenda_ceo": {
    "decidir_hoy": [
      {"texto": "accion concreta -- numero en juego", "empresa": "PAUNO|AMAUTA|NEOPACK", "responsable": "area"}
    ],
    "escalar_semana": [
      {"texto": "seguimiento o reunion necesaria", "empresa": "...", "responsable": "..."}
    ],
    "monitorear": [
      {"texto": "tendencia a vigilar -- indicador", "empresa": "..."}
    ]
  }
}"""

REPORTE_PROMPT = {
    "cuentas_por_cobrar": """Analiza esta imagen de Cuentas por Cobrar y responde SOLO con JSON valido:
{
  "alerta": "una frase de alerta si hay problema, null si todo ok",
  "estado": "red|yellow|green",
  "kpis": [
    {"label": "Morosidad", "valor": "XX%", "meta": "<15%", "estado": "red|yellow|green", "tendencia": "sube|baja|estable"},
    {"label": "CxC total", "valor": "S/X.XM", "meta": null, "estado": "green", "tendencia": "estable"},
    {"label": "Rotacion", "valor": "XXd", "meta": null, "estado": "green", "tendencia": "estable"}
  ],
  "tramos": [
    {"label": "Por vencer", "valor": "S/X.XM", "pct": "XX%", "estado": "green"},
    {"label": "0-30 dias", "valor": "S/XXXK", "pct": "XX%", "estado": "yellow"},
    {"label": "+30 dias", "valor": "S/XXXK", "pct": "XX%", "estado": "red"}
  ],
  "riesgos": [
    {"nombre": "nombre cliente/vendedor", "canal": "canal", "total": "S/XXXK", "vencido30": "S/XXXK", "estado": "red|yellow"}
  ],
  "tendencia_mora": [
    {"mes": "Ene", "pct": 11}, {"mes": "Feb", "pct": 11}
  ],
  "acciones": ["accion 1 concreta", "accion 2", "accion 3"]
}""",

    "cuentas_por_pagar": """Analiza esta imagen de Cuentas por Pagar y responde SOLO con JSON valido:
{
  "alerta": "frase de alerta si hay problema, null si ok",
  "estado": "red|yellow|green",
  "kpis": [
    {"label": "Dias CxP", "valor": "XXXd", "meta": "<90d", "estado": "red|yellow|green", "tendencia": "sube|baja|estable"},
    {"label": "CxP total", "valor": "S/XX.XM", "meta": null, "estado": "yellow", "tendencia": "sube"},
    {"label": "Refinanciado", "valor": "S/X.XM", "meta": null, "estado": "yellow", "tendencia": "estable"}
  ],
  "tramos": [
    {"label": "Vigente", "valor": "S/X.XM", "pct": "XX%", "estado": "green"},
    {"label": "Venc. <15d", "valor": "S/X.XM", "pct": "XX%", "estado": "yellow"},
    {"label": "Venc. 16-90d", "valor": "S/X.XM", "pct": "XX%", "estado": "red"}
  ],
  "proveedores_criticos": [
    {"nombre": "nombre proveedor", "categoria": "categoria", "monto": "S/X.XM", "estado": "red|yellow"}
  ],
  "tendencia_dias": [
    {"mes": "Ago", "dias": 42}
  ],
  "acciones": ["accion 1", "accion 2", "accion 3"]
}""",

    "margen_variable": """Analiza esta imagen de Margen Variable y responde SOLO con JSON valido:
{
  "alerta": "frase de alerta si margen < 46%, null si ok",
  "estado": "red|yellow|green",
  "kpis": [
    {"label": "Margen total", "valor": "XX%", "meta": "52%", "estado": "red|yellow|green", "tendencia": "sube|baja|estable"},
    {"label": "Ventas mes", "valor": "S/X.XM", "meta": null, "estado": "green", "tendencia": "estable"},
    {"label": "Precio/kg", "valor": "S/X.XX", "meta": null, "estado": "green", "tendencia": "estable"}
  ],
  "por_uen": [
    {"uen": "B&D", "margen": "XX%", "estado": "red|yellow|green"},
    {"uen": "MAQUILA", "margen": "XX%", "estado": "red|yellow|green"},
    {"uen": "TIGO", "margen": "XX%", "estado": "red|yellow|green"}
  ],
  "tendencia_margen": [
    {"mes": "Ene", "pct": 52.85}
  ],
  "acciones": ["accion 1", "accion 2"]
}""",

    "mermas": """Analiza esta imagen de Mermas y responde SOLO con JSON valido:
{
  "alerta": "frase si alguna planta > 3%, null si ok",
  "estado": "red|yellow|green",
  "kpis": [
    {"label": "Merma Ate", "valor": "X.X%", "meta": "2%", "estado": "red|yellow|green", "tendencia": "sube|baja|estable"},
    {"label": "Merma Pachacamac", "valor": "X.X%", "meta": "2%", "estado": "red|yellow|green", "tendencia": "sube|baja|estable"},
    {"label": "Merma Terceros", "valor": "X.X%", "meta": "2%", "estado": "red|yellow|green", "tendencia": "sube|baja|estable"}
  ],
  "tendencia_ate": [{"mes": "Ene", "pct": 2.21}],
  "tendencia_pachacamac": [{"mes": "Ene", "pct": 1.59}],
  "acciones": ["accion 1", "accion 2"]
}""",

    "compras": """Analiza esta imagen de Compras (Ratio Consumo/Compra) y responde SOLO con JSON valido.
REGLA: ratio 80-100% = normal (green), >100% = jala stock (yellow si <120%, red si >120%), <80% = sobre-compra (yellow).
El ratio de agosto puede ser alto si el mes esta incompleto, considera eso.
{
  "alerta": "frase si ratio extremo, null si normal",
  "estado": "red|yellow|green",
  "kpis": [
    {"label": "Ratio agosto", "valor": "XXX%", "meta": "80-100%", "estado": "red|yellow|green", "tendencia": "sube|baja|estable"},
    {"label": "Ratio julio", "valor": "XX%", "meta": "80-100%", "estado": "green", "tendencia": "estable"},
    {"label": "Ratio promedio", "valor": "XX%", "meta": "80-100%", "estado": "green", "tendencia": "estable"}
  ],
  "por_categoria": [
    {"categoria": "Materia Prima", "ratio": "XX%", "interpretacion": "normal|jala stock|sobre-compra", "estado": "red|yellow|green"},
    {"categoria": "Envases y Embalajes", "ratio": "XX%", "interpretacion": "...", "estado": "..."},
    {"categoria": "Suministros", "ratio": "XX%", "interpretacion": "...", "estado": "..."},
    {"categoria": "Repuestos", "ratio": "XX%", "interpretacion": "...", "estado": "..."}
  ],
  "tendencia_ratio": [{"mes": "Ago", "ratio": 92.45}],
  "acciones": ["accion 1", "accion 2"]
}""",

    "control_interno": """Analiza esta imagen de Control Interno y responde SOLO con JSON valido:
{
  "alerta": "frase si cumplimiento < 70%, null si ok",
  "estado": "red|yellow|green",
  "kpis": [
    {"label": "Cumplimiento", "valor": "XX%", "meta": "85%", "estado": "red|yellow|green", "tendencia": "sube|baja|estable"},
    {"label": "Puntos criticos", "valor": "XX", "meta": "0", "estado": "red", "tendencia": "estable"},
    {"label": "Satisfactorio", "valor": "XX", "meta": null, "estado": "green", "tendencia": "estable"}
  ],
  "por_gerente": [
    {"nombre": "nombre", "pct": "XX%", "estado": "red|yellow|green"}
  ],
  "tendencia": [{"mes": "Ene", "pct": 44.44}],
  "acciones": ["accion 1", "accion 2"]
}""",

    "sop_inventario": """Analiza esta imagen de S&OP Inventario y responde SOLO con JSON valido:
{
  "alerta": "frase si dead stock > S/1M, null si ok",
  "estado": "red|yellow|green",
  "kpis": [
    {"label": "Dead stock", "valor": "S/X.XM", "meta": "<S/500K", "estado": "red|yellow|green", "tendencia": "sube|baja|estable"},
    {"label": "% Dead", "valor": "XX%", "meta": "<5%", "estado": "red|yellow|green", "tendencia": "estable"},
    {"label": "Working", "valor": "XX%", "meta": ">70%", "estado": "green", "tendencia": "estable"}
  ],
  "clasificacion": [
    {"tipo": "Working", "monto": "S/X.XM", "pct": "XX%", "estado": "green"},
    {"tipo": "Exceso 1", "monto": "S/XXXK", "pct": "XX%", "estado": "yellow"},
    {"tipo": "Exceso 2", "monto": "S/XXXK", "pct": "XX%", "estado": "yellow"},
    {"tipo": "Dead", "monto": "S/X.XM", "pct": "XX%", "estado": "red"}
  ],
  "dead_por_categoria": [
    {"categoria": "Envases y Embalajes", "monto": "S/XXXK", "pct_dead": "XX%", "estado": "red"}
  ],
  "acciones": ["accion 1", "accion 2"]
}""",

    "fill_rate": """Analiza esta imagen de Fill Rate y responde SOLO con JSON valido:
{
  "alerta": "frase si fill rate < 85%, null si ok",
  "estado": "red|yellow|green",
  "kpis": [
    {"label": "Fill Rate", "valor": "XX%", "meta": "98%", "estado": "red|yellow|green", "tendencia": "sube|baja|estable"},
    {"label": "Brecha", "valor": "-Xpp", "meta": "0", "estado": "red|yellow|green", "tendencia": "estable"},
    {"label": "Clientes afect.", "valor": "XX", "meta": null, "estado": "yellow", "tendencia": "estable"}
  ],
  "acciones": ["accion 1", "accion 2"]
}""",

    "productividad": """Analiza esta imagen de Productividad y responde SOLO con JSON valido:
{
  "alerta": null,
  "estado": "red|yellow|green",
  "kpis": [
    {"label": "S//kg producido", "valor": "S/X.XX", "meta": null, "estado": "green", "tendencia": "sube|baja|estable"},
    {"label": "vs 2025", "valor": "S/X.XX", "meta": null, "estado": "green", "tendencia": "baja"},
    {"label": "Planilla total", "valor": "S/X.XM", "meta": null, "estado": "green", "tendencia": "estable"}
  ],
  "tendencia": [{"mes": "Ene", "val": 0.51}],
  "acciones": ["accion 1"]
}""",

    "margen_variable_pag2": """Analiza esta imagen de Avance de Ventas vs Presupuesto y responde SOLO con JSON valido:
{
  "alerta": "frase si avance < 70% a mitad de mes, null si ok",
  "estado": "red|yellow|green",
  "kpis": [
    {"label": "Avance total", "valor": "XX%", "meta": "proporcional al dia", "estado": "red|yellow|green", "tendencia": "sube|baja|estable"},
    {"label": "Ventas reales", "valor": "S/XXXK", "meta": null, "estado": "green", "tendencia": "estable"},
    {"label": "Dia del mes", "valor": "D/31", "meta": null, "estado": "green", "tendencia": "estable"}
  ],
  "por_canal": [
    {"canal": "Moderno", "avance": "XX%", "real": "S/XXXK", "estado": "red|yellow|green"},
    {"canal": "Tradicional", "avance": "XX%", "real": "S/XXXK", "estado": "red|yellow|green"},
    {"canal": "Horeca", "avance": "XX%", "real": "S/XXXK", "estado": "red|yellow|green"}
  ],
  "por_marca": [
    {"marca": "MAQUILA", "avance": "XX%", "real": "S/XXXK", "estado": "green"},
    {"marca": "B&D", "avance": "XX%", "real": "S/XXXK", "estado": "yellow"},
    {"marca": "TIGO", "avance": "XX%", "real": "S/XXXK", "estado": "red"}
  ],
  "acciones": ["accion 1", "accion 2"]
}""",

    "inventario": """Analiza esta imagen de Inventario y responde SOLO con JSON valido:
{
  "alerta": "frase si dead > 5% del total, null si ok",
  "estado": "red|yellow|green",
  "kpis": [
    {"label": "Stock total", "valor": "S/X.XM", "meta": null, "estado": "green", "tendencia": "estable"},
    {"label": "Cobertura", "valor": "XXd", "meta": null, "estado": "green", "tendencia": "estable"},
    {"label": "Dead stock", "valor": "X%", "meta": "<5%", "estado": "red|yellow|green", "tendencia": "estable"}
  ],
  "acciones": ["accion 1"]
}""",

    "ventas": """Analiza esta imagen de Ventas y responde SOLO con JSON valido:
{
  "alerta": "frase si avance < 60% a mitad de mes, null si ok",
  "estado": "red|yellow|green",
  "kpis": [
    {"label": "Avance", "valor": "XX%", "meta": "100%", "estado": "red|yellow|green", "tendencia": "sube|baja|estable"},
    {"label": "Facturacion", "valor": "S/XXXK", "meta": null, "estado": "green", "tendencia": "estable"},
    {"label": "Categoria lider", "valor": "categoria", "meta": null, "estado": "green", "tendencia": "estable"}
  ],
  "acciones": ["accion 1"]
}"""
}


def img_b64(path: Path) -> str:
    return base64.standard_b64encode(path.read_bytes()).decode()


def llamar_claude(client, content: list, max_tokens=1500) -> dict:
    for intento in range(3):
        try:
            resp = client.messages.create(
                model="claude-sonnet-4-5",
                max_tokens=max_tokens,
                messages=[{"role": "user", "content": content}]
            )
            texto = resp.content[0].text.strip()
            if texto.startswith("```"):
                texto = texto.split("```")[1]
                if texto.startswith("json"):
                    texto = texto[4:]
            return json.loads(texto)
        except Exception as e:
            print(f"  Error intento {intento+1}: {e}")
            if intento == 2:
                return {}
    return {}


def generar_semaforo_y_agenda(client, data_dir: Path) -> dict:
    print("Generando semaforo global y agenda CEO...")
    content = []
    for empresa, reportes in REPORTES.items():
        for archivo, _ in reportes[:3]:
            img_path = data_dir / empresa / archivo
            if img_path.exists():
                content.append({"type": "text", "text": f"--- {empresa} / {archivo} ---"})
                content.append({"type": "image", "source": {"type": "base64", "media_type": "image/png", "data": img_b64(img_path)}})
    content.append({"type": "text", "text": SEMAFORO_PROMPT})
    return llamar_claude(client, content, max_tokens=2000)


def generar_reporte(client, empresa: str, archivo: str, tipo: str, data_dir: Path) -> dict:
    img_path = data_dir / empresa / archivo
    if not img_path.exists():
        print(f"  NO ENCONTRADO: {img_path}")
        return {}
    print(f"  Analizando {empresa}/{archivo}...")
    prompt = REPORTE_PROMPT.get(tipo, REPORTE_PROMPT["cuentas_por_cobrar"])
    content = [
        {"type": "text", "text": f"Empresa: {empresa} | Reporte: {tipo}"},
        {"type": "image", "source": {"type": "base64", "media_type": "image/png", "data": img_b64(img_path)}},
        {"type": "text", "text": prompt}
    ]
    return llamar_claude(client, content, max_tokens=1200)


def main():
    if not ANTHROPIC_API_KEY:
        print("ERROR: ANTHROPIC_API_KEY no configurada")
        sys.exit(1)

    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    global_data = generar_semaforo_y_agenda(client, DATA_DIR)

    empresas_data = {}
    for empresa, reportes in REPORTES.items():
        print(f"\nAnalizando {empresa}...")
        empresas_data[empresa] = {"reportes": {}}
        for archivo, tipo in reportes:
            reporte_data = generar_reporte(client, empresa, archivo, tipo, DATA_DIR)
            empresas_data[empresa]["reportes"][tipo] = reporte_data

    resultado = {
        **global_data,
        "generado_en": datetime.now().isoformat(),
        "empresas": empresas_data
    }

    out = DATA_DIR / "summaries.json"
    out.write_text(json.dumps(resultado, ensure_ascii=False, indent=2))
    print(f"\nSummaries guardados en {out}")


if __name__ == "__main__":
    main()
