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

# Mes anterior completo (para filtros DAX — septiembre sin datos completos)
_today = datetime.date.today()
PREV_MONTH = _today.month - 1 if _today.month > 1 else 12
PREV_YEAR  = _today.year  if _today.month > 1 else _today.year - 1

# Tabla/columna de fecha por dataset (descubierta con discover_measures.py #8)
DATE_CONTEXT = {
    "margen":    ("Calendario", "Date"),
    "mermas":    ("Calendario", "Date"),
    "compras":   ("Calendario", "Date"),
    "fill_rate": ("Calendario", "Date"),
    # productividad_ds NO usa este mecanismo genérico — tiene su propio filtro exacto
    # de 9 dimensiones capturado con "Copiar consulta" (ver dax_productividad_pauno).
    # Aplicar aquí YEAR(Date)/MONTH(Date) sería insuficiente: el reporte también
    # filtra por Calendario[Año] (columna distinta a [Date]), MesActual, clasificación
    # contable y exclusiones de producto — confirmado que sin esos filtros el resultado
    # es hasta 38,000x distinto del real (ver commit 2026-09-03).
    "consumo": ("Calendario", "Date"),
}

def dax_productividad_pauno(token, ws_id, dataset_id, measure_name, label="prod"):
    """Ejecuta una medida del dataset '13. Productividad por Funcionario' replicando
    EXACTAMENTE los filtros de página/informe que Power BI aplica en el visual real,
    a nivel PAUNO (sin filtro de Planta). Filtros capturados el 2026-09-03 con
    Analizador de rendimiento > Copiar consulta sobre la tarjeta 'Planilla Total (S/.)'
    con año=2026 (sin mes seleccionado) — validado contra la tarjeta real (S/4,484,364).
    Cualquier cambio en los filtros del reporte en Power BI Desktop debe reflejarse aquí.
    """
    q = f"""EVALUATE
ROW(
  "v",
  CALCULATE(
    [{measure_name}],
    TREATAS({{2026}}, 'Calendario'[Año]),
    TREATAS({{"Otros"}}, 'Calendario'[MesActual]),
    TREATAS({{"AMBAS","GASTO DE PERSONAL OPERATIVO"}}, 'Exl Cuenta Contables'[CLASIFICACION]),
    TREATAS({{"GASTO DE PERSONAL"}}, 'Exl Cuenta Contables'[GASTO_PERSONAL]),
    TREATAS({{"Gasto de Personal Operativo"}}, 'Exl Cuenta Contables'[SUB_CATEGORIA]),
    FILTER(
      KEEPFILTERS(VALUES('Maestra de Facturacion (Total)'[categoria_producto])),
      NOT('Maestra de Facturacion (Total)'[categoria_producto] IN
        {{"BONIFICACION Y REBATES","CHATARRA","INTERESES","MATERIA PRIMA","SERVICIOS","SUMINISTROS",BLANK()}})
    ),
    FILTER(
      KEEPFILTERS(VALUES('Maestra de Facturacion (Total)'[producto])),
      NOT('Maestra de Facturacion (Total)'[producto] IN
        {{"PAVO C/M C/ASA EP CONG (8 KG)","ALIMENTACION COMERCIAL"}})
    ),
    FILTER(
      KEEPFILTERS(VALUES('Maestra de Facturacion (Total)'[estado])),
      NOT('Maestra de Facturacion (Total)'[estado] IN {{"Cancelado"}})
    ),
    FILTER(
      KEEPFILTERS(VALUES('Calendario'[Date])),
      'Calendario'[Date] >= (DATE(2025, 7, 31) + TIME(0, 0, 1))
    )
  )
)"""
    rows = dax(token, ws_id, dataset_id, q, label)
    if rows:
        return rows[0].get("[v]") or rows[0].get("v")
    return None


def dax_productividad_venta_neta_kg(token, ws_id, dataset_id, label="prod_ventaneta"):
    """Venta Neta (KG) del reporte '13. Productividad por Funcionario' — PAUNO.

    No es una medida DAX: en el panel aparece como 'Suma de Peso total KG' porque
    es la COLUMNA 'Maestra de Facturacion (Total)'[Peso total K] con agregación SUM
    directa. Confirmado con Analizador de rendimiento > Copiar consulta sobre la
    tarjeta (2026-09-03) — validado contra la tarjeta real (9,834,538 con año=2026).

    Lleva los mismos 9 filtros que dax_productividad_pauno() más 2 exclusivos de
    esta tarjeta: CATEGORIZACION="VENTA BRUTA" y TIPO DE NEGOCIO N2 no vacío.
    """
    q = f"""EVALUATE
ROW(
  "v",
  CALCULATE(
    SUM('Maestra de Facturacion (Total)'[Peso total K]),
    TREATAS({{"VENTA BRUTA"}}, 'Maestra de Facturacion (Total)'[CATEGORIZACION]),
    FILTER(
      KEEPFILTERS(VALUES('TIPO DE NEGOCIO'[TIPO DE NEGOCIO N2])),
      NOT('TIPO DE NEGOCIO'[TIPO DE NEGOCIO N2] IN {{BLANK()}})
    ),
    TREATAS({{2026}}, 'Calendario'[Año]),
    TREATAS({{"Otros"}}, 'Calendario'[MesActual]),
    TREATAS({{"AMBAS","GASTO DE PERSONAL OPERATIVO"}}, 'Exl Cuenta Contables'[CLASIFICACION]),
    TREATAS({{"GASTO DE PERSONAL"}}, 'Exl Cuenta Contables'[GASTO_PERSONAL]),
    TREATAS({{"Gasto de Personal Operativo"}}, 'Exl Cuenta Contables'[SUB_CATEGORIA]),
    FILTER(
      KEEPFILTERS(VALUES('Maestra de Facturacion (Total)'[categoria_producto])),
      NOT('Maestra de Facturacion (Total)'[categoria_producto] IN
        {{"BONIFICACION Y REBATES","CHATARRA","INTERESES","MATERIA PRIMA","SERVICIOS","SUMINISTROS",BLANK()}})
    ),
    FILTER(
      KEEPFILTERS(VALUES('Maestra de Facturacion (Total)'[producto])),
      NOT('Maestra de Facturacion (Total)'[producto] IN
        {{"PAVO C/M C/ASA EP CONG (8 KG)","ALIMENTACION COMERCIAL"}})
    ),
    FILTER(
      KEEPFILTERS(VALUES('Maestra de Facturacion (Total)'[estado])),
      NOT('Maestra de Facturacion (Total)'[estado] IN {{"Cancelado"}})
    ),
    FILTER(
      KEEPFILTERS(VALUES('Calendario'[Date])),
      'Calendario'[Date] >= (DATE(2025, 7, 31) + TIME(0, 0, 1))
    )
  )
)"""
    rows = dax(token, ws_id, dataset_id, q, label)
    if rows:
        return rows[0].get("[v]") or rows[0].get("v")
    return None


def _dax_consumo_filtro_kardex(token, ws_id, dataset_id, value_expr, label, anio=None):
    """Aplica los filtros confirmados con Copiar consulta que comparten las
    tarjetas 'Costo Total' y 'Producción Neta (KG)' del reporte '14. Consumo
    Materiales indirectos de producción'.

    2026-09-03: primera captura, SIN año seleccionado en el reporte -> sin filtro
    de fecha, acumulado histórico total. 2026-09-04: se repitió la captura con
    el selector "2026" marcado en el reporte -> agrega TREATAS({{2026}},
    'Calendario'[Año]). Por eso `anio` es parámetro: pasar PREV_YEAR (o el año
    que corresponda) para obtener el dato del año, u omitir para el histórico total.
    `value_expr` es la expresión DAX del valor (columna con SUM o [Medida]).
    """
    # Se construye por concatenación simple (no f-string para todo el bloque) para
    # no arriesgar un error de escapado de llaves entre DAX y Python.
    filtro_anio = ("TREATAS({" + str(int(anio)) + "}, 'Calendario'[Año]),\n    ") if anio else ""
    q = (
        'EVALUATE\nROW(\n  "v",\n  CALCULATE(\n    ' + value_expr + ",\n    " +
        filtro_anio +
        "TREATAS({\"Costo\"}, 'PLANTA POR CECOS'[TIPO DE OPERACION]),\n"
        "    FILTER(\n"
        "      KEEPFILTERS(VALUES('Maestra de Kardex (Total)'[categoria_hijo])),\n"
        "      NOT('Maestra de Kardex (Total)'[categoria_hijo] IN {\"ACUERDOS COMERCIALES\"})\n"
        "    ),\n"
        "    FILTER(\n"
        "      KEEPFILTERS(VALUES('Maestra de Kardex (Total)'[CUENTA ORIGEN])),\n"
        "      NOT('Maestra de Kardex (Total)'[CUENTA ORIGEN] IN {BLANK()})\n"
        "    ),\n"
        "    FILTER(\n"
        "      KEEPFILTERS(VALUES('Maestra de Kardex (Total)'[cuenta_analitica])),\n"
        "      NOT('Maestra de Kardex (Total)'[cuenta_analitica] IN\n"
        "        {BLANK(),\"[941002] CONTROL INTERNO\",\"ALMACEN ATE\",\"ALMACEN PACHACAMAC\"})\n"
        "    )\n"
        "  )\n"
        ")"
    )
    rows = dax(token, ws_id, dataset_id, q, label)
    if rows:
        return rows[0].get("[v]") or rows[0].get("v")
    return None


def dax_consumo_costo_total(token, ws_id, dataset_id, label="consumo_costo"):
    """'Costo Total' del reporte '14. Consumo Materiales indirectos de producción'.
    Es la medida 'Maestra de Kardex (Total)'[Costo total validado].
    Filtrado por Calendario[Año]=PREV_YEAR desde 2026-09-04 (antes era histórico
    total sin año — se corrigió tras confirmar con una segunda Copiar consulta,
    esta vez con el selector de año marcado en el reporte).
    """
    return _dax_consumo_filtro_kardex(token, ws_id, dataset_id, "[Costo total validado]", label, anio=PREV_YEAR)


def dax_consumo_produccion_neta_kg(token, ws_id, dataset_id, label="consumo_prodneta"):
    """'Producción Neta (KG)' del reporte '14. Consumo Materiales indirectos de producción'.
    Es la medida 'Medidas'[Producción (KG) Odoo] — mismos filtros que Costo Total,
    incluido Calendario[Año]=PREV_YEAR (ver dax_consumo_costo_total).
    """
    return _dax_consumo_filtro_kardex(token, ws_id, dataset_id, "[Producción (KG) Odoo]", label, anio=PREV_YEAR)


def _dax_consumo_filtro_venta(token, ws_id, dataset_id, value_expr, label):
    """Aplica los 7 filtros confirmados con Copiar consulta que comparten las
    tarjetas 'Venta Neta (KG)' y 'Costo x TN Vendida' del reporte '14. Consumo
    Materiales indirectos de producción' (2026-09-03): filtro de fecha desde
    2025-07-01, CATEGORIZACION='VENTA BRUTA', TIPO DE NEGOCIO N2 no vacío, y los
    mismos 4 de Planta/Kardex que usa 'Costo Total' en este mismo reporte.
    `value_expr` es la expresión DAX del valor (columna con SUM o [Medida]).
    """
    q = f"""EVALUATE
ROW(
  "v",
  CALCULATE(
    {value_expr},
    FILTER(
      KEEPFILTERS(VALUES('Calendario'[Date])),
      'Calendario'[Date] >= (DATE(2025, 7, 1) + TIME(0, 0, 1))
    ),
    TREATAS({{"VENTA BRUTA"}}, 'Maestra de Facturacion (Total)'[CATEGORIZACION]),
    FILTER(
      KEEPFILTERS(VALUES('TIPO DE NEGOCIO'[TIPO DE NEGOCIO N2])),
      NOT('TIPO DE NEGOCIO'[TIPO DE NEGOCIO N2] IN {{BLANK()}})
    ),
    TREATAS({{"Costo"}}, 'PLANTA POR CECOS'[TIPO DE OPERACION]),
    FILTER(
      KEEPFILTERS(VALUES('Maestra de Kardex (Total)'[categoria_hijo])),
      NOT('Maestra de Kardex (Total)'[categoria_hijo] IN {{"ACUERDOS COMERCIALES"}})
    ),
    FILTER(
      KEEPFILTERS(VALUES('Maestra de Kardex (Total)'[CUENTA ORIGEN])),
      NOT('Maestra de Kardex (Total)'[CUENTA ORIGEN] IN {{BLANK()}})
    ),
    FILTER(
      KEEPFILTERS(VALUES('Maestra de Kardex (Total)'[cuenta_analitica])),
      NOT('Maestra de Kardex (Total)'[cuenta_analitica] IN
        {{BLANK(),"[941002] CONTROL INTERNO","ALMACEN ATE","ALMACEN PACHACAMAC"}})
    )
  )
)"""
    rows = dax(token, ws_id, dataset_id, q, label)
    if rows:
        return rows[0].get("[v]") or rows[0].get("v")
    return None


def dax_consumo_venta_neta_kg(token, ws_id, dataset_id, label="consumo_ventaneta"):
    """'Venta Neta (KG)' del reporte '14. Consumo Materiales indirectos de producción'.
    No es medida, es la columna 'Maestra de Facturacion (Total)'[Peso total KG]
    (con G — distinta de [Peso total K] sin G que usa Productividad) con SUM directo.
    """
    return _dax_consumo_filtro_venta(
        token, ws_id, dataset_id,
        "SUM('Maestra de Facturacion (Total)'[Peso total KG])", label)


def dax_consumo_costo_x_tn_vendida(token, ws_id, dataset_id, label="consumo_costotnvend"):
    """'Costo x TN Vendida' del reporte '14. Consumo Materiales indirectos de producción'.
    Es la medida 'Maestra de Kardex (Total)'[Ratio costo / kg] — mismos 7 filtros
    que Venta Neta (KG), confirmados con Copiar consulta el 2026-09-03.
    """
    return _dax_consumo_filtro_venta(token, ws_id, dataset_id, "[Ratio costo / kg]", label)


def dax_consumo_costo_x_tn_producida(token, ws_id, dataset_id, label="consumo_costotnprod"):
    """'Costo x TN Producida' del reporte '14. Consumo Materiales indirectos de producción'.
    Es la medida 'Medidas'[Costo x ton producida] — mismos 7 filtros que
    Venta Neta (KG) / Costo x TN Vendida, confirmados con Copiar consulta el 2026-09-03.
    """
    return _dax_consumo_filtro_venta(token, ws_id, dataset_id, "[Costo x ton producida]", label)


def _dax_margen_filtros(uen, anio):
    """Los 12 filtros de página confirmados con Copiar consulta (2026-09-04) sobre
    la tabla "Análisis de Ventas" del reporte de Margen: CATEGORIZACION='VENTA BRUTA',
    exclusión de "BONIFICACION SELL IN", Calendario[Año]=<anio>, y 8 filtros más de
    exclusión de categoria_producto/producto/estado/CUENTA ORIGEN sobre 4 tablas
    ('Exl A Maestra de Facturas de Venta', 'Maestra de Facturacion (Total)',
    'Maestra de Kardex (Total)', 'OrdenxFactura').

    Si `uen` no es None, agrega un 13er filtro TIPO DE NEGOCIO N1=<uen> — confirmado
    capturando la misma consulta con la tarjeta filtrada a "B&D". Sin `uen` es el
    consolidado total (confirmado con una segunda captura sin ese filtro).
    """
    filtro_uen = ""
    if uen:
        uen_esc = uen.replace('"', '\\"')
        filtro_uen = "    TREATAS({\"" + uen_esc + "\"}, 'Exl Tipo de Negocio'[TIPO DE NEGOCIO N1]),\n"
    return (
        "TREATAS({\"VENTA BRUTA\"}, 'Exl A Maestra de Facturas de Venta'[CATEGORIZACION]),\n"
        "    FILTER(\n"
        "      KEEPFILTERS(VALUES('Exl A Maestra de Facturas de Venta'[DESCRIPCION])),\n"
        "      NOT('Exl A Maestra de Facturas de Venta'[DESCRIPCION] IN {\"BONIFICACION SELL IN\"})\n"
        "    ),\n"
        + filtro_uen +
        "    TREATAS({" + str(int(anio)) + "}, 'Calendario'[Año]),\n"
        "    FILTER(\n"
        "      KEEPFILTERS(VALUES('Exl A Maestra de Facturas de Venta'[categoria_producto])),\n"
        "      NOT('Exl A Maestra de Facturas de Venta'[categoria_producto] IN {\"CHATARRA\",\"SERVICIOS\",BLANK()})\n"
        "    ),\n"
        "    FILTER(\n"
        "      KEEPFILTERS(VALUES('Exl A Maestra de Facturas de Venta'[producto])),\n"
        "      NOT('Exl A Maestra de Facturas de Venta'[producto] IN\n"
        "        {\"PAVO C/M C/ASA EP CONG (8 KG)\",\"ALIMENTACION COMERCIAL\"})\n"
        "    ),\n"
        "    FILTER(\n"
        "      KEEPFILTERS(VALUES('Exl A Maestra de Facturas de Venta'[estado])),\n"
        "      NOT('Exl A Maestra de Facturas de Venta'[estado] IN {\"Cancelado\"})\n"
        "    ),\n"
        "    FILTER(\n"
        "      KEEPFILTERS(VALUES('Maestra de Facturacion (Total)'[categoria_producto])),\n"
        "      NOT('Maestra de Facturacion (Total)'[categoria_producto] IN\n"
        "        {BLANK(),\"BONIFICACION Y REBATES\",\"INTERESES\",\"MATERIA PRIMA\",\"SUMINISTROS\",\"SERVICIOS\"})\n"
        "    ),\n"
        "    FILTER(\n"
        "      KEEPFILTERS(VALUES('Maestra de Facturacion (Total)'[estado])),\n"
        "      NOT('Maestra de Facturacion (Total)'[estado] IN {\"Cancelado\"})\n"
        "    ),\n"
        "    FILTER(\n"
        "      KEEPFILTERS(VALUES('Maestra de Kardex (Total)'[CUENTA ORIGEN])),\n"
        "      NOT('Maestra de Kardex (Total)'[CUENTA ORIGEN] IN {BLANK()})\n"
        "    ),\n"
        "    FILTER(\n"
        "      KEEPFILTERS(VALUES('OrdenxFactura'[categoria_producto])),\n"
        "      NOT('OrdenxFactura'[categoria_producto] IN\n"
        "        {BLANK(),\"BONIFICACION Y REBATES\",\"CHATARRA\",\"SERVICIOS\"})\n"
        "    ),\n"
        "    FILTER(\n"
        "      KEEPFILTERS(VALUES('OrdenxFactura'[estado])),\n"
        "      NOT('OrdenxFactura'[estado] IN {\"Cancelado\"})\n"
        "    ),\n"
        "    FILTER(\n"
        "      KEEPFILTERS(VALUES('OrdenxFactura'[producto])),\n"
        "      NOT('OrdenxFactura'[producto] IN\n"
        "        {\"PAVO C/M C/ASA EP CONG (8 KG)\",\"ALIMENTACION COMERCIAL\"})\n"
        "    )"
    )


def _dax_margen_precio_costo(token, ws_id, dataset_id, uen, anio, label):
    """Ejecuta Precio x Kilo / Costo x Kilo con los filtros de _dax_margen_filtros.
    Las medidas son 'Medidas Julito'[Precio x Kilo] y [Costo x Kilo] — mismo nombre
    que las medidas ya validadas para el margen consolidado, aquí referenciadas sin
    prefijo de tabla porque el nombre de medida es único en el modelo.
    """
    filtros = _dax_margen_filtros(uen, anio)
    q = (
        'EVALUATE\nROW(\n  "precio", CALCULATE([Precio x Kilo],\n    ' + filtros + '\n  ),\n'
        '  "costo", CALCULATE([Costo x Kilo],\n    ' + filtros + "\n  )\n)"
    )
    rows = dax(token, ws_id, dataset_id, q, label)
    if rows:
        precio = to_float(rows[0].get("[precio]") or rows[0].get("precio"))
        costo = to_float(rows[0].get("[costo]") or rows[0].get("costo"))
        return precio, costo
    return None, None


def dax_margen_uen(token, ws_id, dataset_id, uen, anio, label="margen_uen"):
    """Precio/kg y Costo/kg filtrados por UEN (TIPO DE NEGOCIO N1 = "B&D" / "MAQUILA" / "TIGO").
    La consulta original de Power BI agrupa por mes (visual de tabla); aquí se
    colapsa a un solo agregado del año completo.
    """
    return _dax_margen_precio_costo(token, ws_id, dataset_id, uen, anio, label)


def dax_margen_total(token, ws_id, dataset_id, anio, label="margen_total"):
    """Precio/kg y Costo/kg consolidados (todas las UEN), con el mismo filtro de
    página de 12 condiciones que dax_margen_uen pero sin el filtro de UEN —
    confirmado con una segunda Copiar consulta capturada sin ese filtro activo
    (2026-09-04). Reemplaza al valor que traía el escaneo genérico con solo
    filtro de fecha, que no aplicaba las 11 exclusiones adicionales del reporte.
    """
    return _dax_margen_precio_costo(token, ws_id, dataset_id, None, anio, label)


def dax_mermas_uen(token, ws_id, dataset_id, medida, anio, extra_filtro=None, label="mermas_uen"):
    """Ejecuta una medida de 'Tabla Mermas' (ej. '% Merma total B&D') del reporte
    '4. Reporte de mermas' con el filtro de página común, confirmado con Copiar
    consulta el 2026-09-04 sobre las tarjetas 'Merma B&D' y 'Merma TIGO':
      Calendario[Date] >= 2025-07-31 (mismo corte que otros reportes)
      Calendario[Año] = <anio>
      Maestra de Facturacion (Total)[categoria_producto] excluye 7 categorías
      Maestra de Facturacion (Total)[producto] excluye 2 productos
      Maestra de Facturacion (Total)[estado] excluye "Cancelado"

    `medida` es el nombre exacto de la medida en 'Tabla Mermas', ej.
    "% Merma total B&D" / "% Merma total TIGO".

    `extra_filtro` es un filtro DAX adicional propio de la tarjeta (string, sin
    coma final) — TIGO trae uno que B&D no tiene:
    'Tabla Mermas'[TIPO DE BASE] no vacío. No se asume que todas las UEN
    comparten exactamente los mismos filtros — cada una se confirma por separado.
    """
    medida_esc = medida.replace('"', '\\"')
    extra = f"    {extra_filtro},\n" if extra_filtro else ""
    q = (
        'EVALUATE\nROW(\n  "v", CALCULATE(\n    ' + f"'Tabla Mermas'[{medida_esc}]" + ',\n'
        "    FILTER(\n"
        "      KEEPFILTERS(VALUES('Calendario'[Date])),\n"
        "      'Calendario'[Date] >= (DATE(2025, 7, 31) + TIME(0, 0, 1))\n"
        "    ),\n"
        + extra +
        "    TREATAS({" + str(int(anio)) + "}, 'Calendario'[Año]),\n"
        "    FILTER(\n"
        "      KEEPFILTERS(VALUES('Maestra de Facturacion (Total)'[categoria_producto])),\n"
        "      NOT('Maestra de Facturacion (Total)'[categoria_producto] IN\n"
        "        {\"BONIFICACION Y REBATES\",\"CHATARRA\",\"INTERESES\",\"MATERIA PRIMA\",\"SERVICIOS\",\"SUMINISTROS\",BLANK()})\n"
        "    ),\n"
        "    FILTER(\n"
        "      KEEPFILTERS(VALUES('Maestra de Facturacion (Total)'[producto])),\n"
        "      NOT('Maestra de Facturacion (Total)'[producto] IN\n"
        "        {\"PAVO C/M C/ASA EP CONG (8 KG)\",\"ALIMENTACION COMERCIAL\"})\n"
        "    ),\n"
        "    FILTER(\n"
        "      KEEPFILTERS(VALUES('Maestra de Facturacion (Total)'[estado])),\n"
        "      NOT('Maestra de Facturacion (Total)'[estado] IN {\"Cancelado\"})\n"
        "    )\n"
        "  )\n"
        ")"
    )
    rows = dax(token, ws_id, dataset_id, q, label)
    if rows:
        return to_float(rows[0].get("[v]") or rows[0].get("v"))
    return None


def dax_mermas_planta(token, ws_id, dataset_id, medida, almacenes, anio, con_tipo_base=True, label="mermas_planta"):
    """Medida de 'Tabla Mermas' del reporte de mermas (pestaña RESUMEN, gráfico
    'Merma Mensual por Planta') filtrada por 'Tabla Mermas'[almacen].

    Confirmado con Copiar consulta el 2026-09-04, capturado directamente del
    panel del Analizador de rendimiento (fila "PLANTA ATE"/"PLANTA PACHACAMAC"/
    "PLANTA TERCEROS" bajo el grupo "MERMA MENSUAL POR PLANTA", no las filas
    duplicadas de "MERMA DIARIA POR PLANTA" más abajo en la misma lista):
      Planta Ate         -> almacen "Lácteos Producción"                 · medida
                             genérica 'Tabla Mermas'[% Merma total] (sin sufijo)
                             · SÍ trae filtro TIPO DE BASE
      Planta Pachacámac  -> almacen "Salsas Producción"                   · medida
                             'Tabla Mermas'[% Merma total SALSAS]
                             · SÍ trae filtro TIPO DE BASE
      Planta Terceros    -> almacenes de maquiladores externos (Abuela
                             Maquila, Piamonte Maquila, Lácteos Dosimetría)
                             · medida 'Tabla Mermas'[% Merma total MAQUILA]
                             (una M — distinta de la medida UEN "MMAQUILA")
                             · NO trae filtro TIPO DE BASE (confirmado ausente
                             en su Copiar consulta, a diferencia de las otras 2)
    Cada planta puede tener su propia medida y sus propios filtros — no se
    asume el mismo patrón para todas, igual que ya pasó con Mermas por UEN.

    `almacenes` es una lista de valores de [almacen] — normalmente uno solo,
    pero TREATAS admite varios si una planta agrupa más de un almacén.
    `con_tipo_base` controla si se incluye el filtro TIPO DE BASE — pasar
    False para Terceros.
    """
    medida_esc = medida.replace('"', '\\"')
    vals = ",".join(f'"{a}"' for a in almacenes)
    tipo_base = (
        "    FILTER(\n"
        "      KEEPFILTERS(VALUES('Tabla Mermas'[TIPO DE BASE])),\n"
        "      NOT('Tabla Mermas'[TIPO DE BASE] IN {BLANK()})\n"
        "    ),\n"
    ) if con_tipo_base else ""
    q = (
        'EVALUATE\nROW(\n  "v", CALCULATE(\n    ' + f"'Tabla Mermas'[{medida_esc}]" + ',\n'
        "    FILTER(\n"
        "      KEEPFILTERS(VALUES('Calendario'[Date])),\n"
        "      'Calendario'[Date] >= (DATE(2025, 7, 31) + TIME(0, 0, 1))\n"
        "    ),\n"
        + tipo_base +
        "    TREATAS({" + vals + "}, 'Tabla Mermas'[almacen]),\n"
        "    TREATAS({" + str(int(anio)) + "}, 'Calendario'[Año]),\n"
        "    FILTER(\n"
        "      KEEPFILTERS(VALUES('Maestra de Facturacion (Total)'[categoria_producto])),\n"
        "      NOT('Maestra de Facturacion (Total)'[categoria_producto] IN\n"
        "        {\"BONIFICACION Y REBATES\",\"CHATARRA\",\"INTERESES\",\"MATERIA PRIMA\",\"SERVICIOS\",\"SUMINISTROS\",BLANK()})\n"
        "    ),\n"
        "    FILTER(\n"
        "      KEEPFILTERS(VALUES('Maestra de Facturacion (Total)'[producto])),\n"
        "      NOT('Maestra de Facturacion (Total)'[producto] IN\n"
        "        {\"PAVO C/M C/ASA EP CONG (8 KG)\",\"ALIMENTACION COMERCIAL\"})\n"
        "    ),\n"
        "    FILTER(\n"
        "      KEEPFILTERS(VALUES('Maestra de Facturacion (Total)'[estado])),\n"
        "      NOT('Maestra de Facturacion (Total)'[estado] IN {\"Cancelado\"})\n"
        "    )\n"
        "  )\n"
        ")"
    )
    rows = dax(token, ws_id, dataset_id, q, label)
    if rows:
        return to_float(rows[0].get("[v]") or rows[0].get("v"))
    return None


def dax_inventario_kardex(token, ws_id, dataset_id, medida, anio, mes, label="inv_kardex"):
    """Medida de 'KARDEX TOTAL' del reporte '6. Rotación de inventario'
    (página "RI (Clasificación)", gráfico "Evolución de Inventario").

    Confirmado con Copiar consulta el 2026-09-04. Filtros propios de este
    dataset — distintos a los de Margen/Mermas, sin las exclusiones de
    categoria_producto/producto/estado que usan esos otros reportes:
      Calendario[Mes Año] <> "Jun 2025" (el reporte excluye ese mes siempre,
        probablemente por dato parcial al inicio del histórico)
      KARDEX TOTAL[Fecha] dentro del mes/año pedido

    `medida` ej. "Días Rotación" (≈ Cobertura Total en el visual) o
    "Días Rotación MP" (≈ Cobertura MP). También sirve para
    "Consumo Acumulado Total S/ saldo".
    """
    medida_esc = medida.replace('"', '\\"')
    anio, mes = int(anio), int(mes)
    anio_sig, mes_sig = (anio + 1, 1) if mes == 12 else (anio, mes + 1)
    q = (
        'EVALUATE\nROW(\n  "v", CALCULATE(\n    ' + f"'KARDEX TOTAL'[{medida_esc}]" + ',\n'
        "    FILTER(\n"
        "      KEEPFILTERS(VALUES('Calendario'[Mes Año])),\n"
        "      NOT('Calendario'[Mes Año] IN {\"Jun 2025\"})\n"
        "    ),\n"
        "    FILTER(\n"
        "      KEEPFILTERS(VALUES('KARDEX TOTAL'[Fecha])),\n"
        "      AND(\n"
        f"        'KARDEX TOTAL'[Fecha] >= DATE({anio}, {mes}, 1),\n"
        f"        'KARDEX TOTAL'[Fecha] < DATE({anio_sig}, {mes_sig}, 1)\n"
        "      )\n"
        "    )\n"
        "  )\n"
        ")"
    )
    rows = dax(token, ws_id, dataset_id, q, label)
    if rows:
        return to_float(rows[0].get("[v]") or rows[0].get("v"))
    return None


def dax_inventario_composicion(token, ws_id, dataset_id, label="inv_composicion"):
    """Desglose de Saldo Soles por 'SALDO ACTUAL'[Clasificación Segun Consumo]
    (WORKING / EXCESO 1 / EXCESO 2 / DEAD) del gráfico "Composición de
    Inventario" en '6. Rotación de inventario'.

    Confirmado con Copiar consulta el 2026-09-04. A diferencia de
    dax_inventario_kardex, esta consulta devuelve VARIAS filas (una por
    categoría) — no un solo valor. El filtro de fecha es amplio
    (2025-06-30 a 2027-01-01): la clasificación es una foto actual, no un
    corte mensual.

    Devuelve una lista de (clasificacion, saldo_soles) ordenada de mayor a
    menor, igual que el gráfico de cascada real.
    """
    q = (
        "EVALUATE\n"
        "SUMMARIZECOLUMNS(\n"
        "  'SALDO ACTUAL'[Clasificación Segun Consumo],\n"
        "  FILTER(\n"
        "    KEEPFILTERS(VALUES('KARDEX TOTAL'[Fecha])),\n"
        "    AND(\n"
        "      'KARDEX TOTAL'[Fecha] >= DATE(2025, 6, 30),\n"
        "      'KARDEX TOTAL'[Fecha] < DATE(2027, 1, 1)\n"
        "    )\n"
        "  ),\n"
        '  "SumSaldo_Soles", CALCULATE(SUM(\'SALDO ACTUAL\'[Saldo Soles]))\n'
        ")\n"
        "ORDER BY [SumSaldo_Soles] DESC"
    )
    rows = dax(token, ws_id, dataset_id, q, label)
    out = []
    for r in rows or []:
        clasif = None
        saldo = None
        for k, v in r.items():
            if "Clasificaci" in k:
                clasif = v
            elif "SumSaldo" in k:
                saldo = to_float(v)
        if clasif is not None:
            out.append((clasif, saldo))
    return out


def dax_control_interno(token, ws_id, dataset_id, medida, empresa="Pauno", label="control_interno"):
    """Medida de 'Pauno Registro de Ejecuciones' del reporte '8. Reporte de
    auditoría' (Dashboard de Control Interno).

    Confirmado con Copiar consulta el 2026-09-04 sobre la tarjeta
    "Satisfactorio": el único filtro es Empresa="Pauno" — SIN filtro de
    fecha ni mes, a pesar de que el dashboard tiene un selector de mes.
    Los KPIs de resumen (Puntos de Control, Satisfactorio, Con Observaciones,
    Crítico, % Cumplimiento) son un acumulado histórico total, no del mes.

    `medida` ej. "Satisfactorio", "Con Observaciones", "Critico",
    "Puntos de Control", "% Cumplimiento".
    """
    medida_esc = medida.replace('"', '\\"')
    empresa_esc = empresa.replace('"', '\\"')
    q = (
        'EVALUATE\nROW(\n  "v", CALCULATE(\n    ' + f"'Pauno Registro de Ejecuciones'[{medida_esc}]" + ',\n'
        f'    TREATAS({{"{empresa_esc}"}}, \'Pauno Registro de Ejecuciones\'[Empresa])\n'
        "  )\n"
        ")"
    )
    rows = dax(token, ws_id, dataset_id, q, label)
    if rows:
        return to_float(rows[0].get("[v]") or rows[0].get("v"))
    return None


def dax_planes_accion(token, ws_id, dataset_id, medida, empresa="Pauno", label="planes_accion"):
    """Medida de 'Planes de Accion' (mismo dataset que Control Interno,
    tabla distinta a 'Pauno Registro de Ejecuciones').

    Confirmado con Copiar consulta el 2026-09-04 sobre la tarjeta
    "Planes Abiertos": filtro único Empresa="Pauno" (mismo patrón que
    Control Interno, sin fecha), pero usando SUMMARIZECOLUMNS + IGNORE
    en vez de CALCULATE (así lo genera Power BI para esta tarjeta).

    `medida` ej. "Planes Abiertos", "Planes Atrasados", "Total Planes de Acción".
    """
    medida_esc = medida.replace('"', '\\"')
    empresa_esc = empresa.replace('"', '\\"')
    alias = "".join(c if c.isalnum() else "_" for c in medida)
    q = (
        "DEFINE VAR __DS0FilterTable = \n"
        f'\tTREATAS({{"{empresa_esc}"}}, \'Pauno Registro de Ejecuciones\'[Empresa])\n\n'
        "EVALUATE\n"
        f'\tSUMMARIZECOLUMNS(__DS0FilterTable, "{alias}", IGNORE(\'Planes de Accion\'[{medida_esc}]))'
    )
    rows = dax(token, ws_id, dataset_id, q, label)
    if rows:
        return to_float(rows[0].get(f"[{alias}]") or rows[0].get(alias))
    return None


def dax_control_interno_sum(token, ws_id, dataset_id, columna, empresa="Pauno", label="control_interno_sum"):
    """Suma de una columna (no medida) de 'Pauno Registro de Ejecuciones'.

    Confirmado con Copiar consulta el 2026-09-04 sobre la columna
    '¿Ejecutado?': mismo filtro Empresa="Pauno" que el resto de Control
    Interno, pero usando SUMMARIZECOLUMNS + IGNORE(CALCULATE(SUM(...)))
    en vez de referenciar una medida ya definida en el modelo.

    `columna` ej. "¿Ejecutado?".
    """
    columna_esc = columna.replace('"', '\\"')
    empresa_esc = empresa.replace('"', '\\"')
    alias = "Sumv_" + "".join(c for c in columna if c.isalnum()) + "_"
    q = (
        "DEFINE VAR __DS0FilterTable = \n"
        f'\tTREATAS({{"{empresa_esc}"}}, \'Pauno Registro de Ejecuciones\'[Empresa])\n\n'
        "EVALUATE\n"
        "\tSUMMARIZECOLUMNS(\n"
        "\t\t__DS0FilterTable,\n"
        f'\t\t"{alias}", IGNORE(CALCULATE(SUM(\'Pauno Registro de Ejecuciones\'[{columna_esc}])))\n'
        "\t)"
    )
    rows = dax(token, ws_id, dataset_id, q, label)
    if rows:
        return to_float(rows[0].get(f"[{alias}]") or rows[0].get(alias))
    return None


def dax_prev_month(token, ws_id, dataset_id, measure_name, date_tbl, date_col, label="dated"):
    """Ejecuta medida filtrada al mes anterior completo."""
    q = f"""EVALUATE
ROW("v",
  CALCULATE(
    [{measure_name}],
    FILTER(ALL('{date_tbl}'), YEAR('{date_tbl}'[{date_col}]) = {PREV_YEAR} && MONTH('{date_tbl}'[{date_col}]) = {PREV_MONTH})
  )
)"""
    rows = dax(token, ws_id, dataset_id, q, label)
    if rows:
        return rows[0].get("[v]") or rows[0].get("v")
    return None

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
        "fill_rate":      "7f4ebe22-5e90-4e35-973b-4af3c58497e5",  # 12. Calculo de Provisiones
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
        # ── CONFIRMADOS por prueba directa ──
        "Cuentas x Pagar", "CUENTAS X PAGAR", "Refinanciamiento", "Proveedores",
        # ── Tramos vencido (deducidos de la captura del reporte) ──
        "Vigente CxP", "CxP Vigente", "No Vencido CxP",
        "Vencido < 15", "Vencido Menor 15d", "Vencido 0-15d",
        "Vencido 16-30d", "Vencido 30d", "Vencido 1-30",
        "Vencido 31-90", "Vencido 31-90d", "Vencido 90d",
        "Vencido >90", "Vencido >90d", "Vencido Mayor 90",
        "Saldo Vigente", "Saldo Vencido",
        # ── Días de pago (eje del gráfico) ──
        "Dias CxP", "DPP", "Dias Pago", "Rotacion CxP",
        "Dias Promedio Pago", "Dias de Pago", "Días de Pago",
        "PPP", "Plazo Pago", "Plazo Promedio Pago",
        # ── Otros candidatos ──
        "Total CxP", "Saldo CxP", "CxP Total",
        "Deuda Total", "Total Deuda", "Monto CxP",
    ],
    "margen": [
        # Columnas exactas visibles en tabla "DETALLE ANALISIS DE COSTOS" (Power BI PAUNO)
        "% Margen", "Costo Total", "P.Unit", "P.Unit NC", "C.Unit",
        # Nombres de tabs visibles en Power BI PAUNO
        "R. MARGEN", "R. Margen", "R MARGEN",
        "R. COSTO UNIT", "R. Costo Unit", "Costo Unitario", "Costo Unit",
        # Margen variable
        "Margen Variable", "% Margen Variable", "MV", "% MV", "MV%", "Margen Bruto",
        "Margen", "Costo Variable", "% Costo Variable",
        "% MV PAUNO", "MV PAUNO", "Margen Variable %", "Pct MV",
        # Ventas
        "Ventas Mes Actual", "Ventas Actuales", "Ventas", "Venta Total", "Venta Neta",
        "Ventas Mes Anterior", "Ventas Anterior",
        # Precio / Costo por kg
        "Precio/kg", "Precio kg", "Precio Promedio kg", "Precio x Kilo", "Precio Kilo",
        "Costo/kg", "Costo kg", "Costo x Kilo", "Costo Kilo",
        # Nombres con mayúsculas/abrev peruanas
        "MARGEN VARIABLE", "% MARGEN", "MARGEN", "MV TOTAL",
        "VENTA NETA", "VENTAS NETAS", "VENTA MES",
        "Ing. Ventas", "Ingresos", "Ingresos Ventas",
        "Resultado Bruto", "Utilidad Bruta", "% Utilidad Bruta",
    ],
    "mermas": [
        # Columnas exactas del "Detalle de Mermas por Planta" (Control de Producción)
        "REAL", "STD", "DESVIACION", "ITEM DESVIADO", "% MERMAS",
        "Stock Prod", "Stock Obs", "Merma",
        # Medidas estándar
        "% Merma", "Merma %", "Tasa Merma", "Merma Total", "% Merma Total",
        "% Merma Ate", "Merma Ate", "Merma Planta Ate",
        "% Merma Pachacamac", "Merma Pachacamac", "Merma Lurin",
        "% Merma Terceros", "Merma Terceros",
        # Por UEN
        "% Merma B&D", "Merma B&D", "% Merma BD",
        "% Merma Tigo", "Merma Tigo", "% Merma TIGO",
        "% Merma Maquila", "Merma Maquila", "% Merma MAQUILA",
        "Merma Valorizada", "Kg Merma",
    ],
    "compras": [
        # Columnas exactas "ANALISIS DE MATERIALES (PLANIFICACION Y COMPRA)" / "Reporte Compras"
        "Lead Time", "Stock PP", "Stock Valorizado",
        "Costo Unit", "Costo Unitario",
        "Cons Mes Ant", "Consumo Mes Anterior",
        "Consumo Prom 3M", "Consumo Promedio 3M",
        "Consumo Prom 6M", "Consumo Promedio 6M",
        "Faltantes", "Requerimiento",
        # Fila total: Consumo | Compra | Ratio por período
        "Ratio", "Cant Compra", "Cant Consumo",
        "% Ratio", "Ratio C/V", "Ratio Compras", "Ratio Consumo/Compra",
        "Eficiencia", "Eficiencia Costo", "% Eficiencia",
        "Ratio C/C", "Consumo/Compra", "C/C",
        "Consumo Total", "Consumo Periodo", "Importe Consumo", "Monto Consumo",
        "Compra Total", "Monto Compras", "Importe Compras", "Valor Compras",
        "MP Consumo", "MP Compras",
        "Total Compras Periodo", "Ordenes de Compra",
        "Compra Mes", "Consumo Mes", "Compras Mes",
        "Valor Compra", "Valor Consumo",
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
        # Nombres REALES confirmados leyendo el panel Datos > Measures en Power BI Service
        # (Editar informe) y validados con EVALUATE ROW(...) en DAX Query View — 2026-09-03.
        "SUM PROD",              # -> Producción Total (KG)
        "GASTO TOTAL",           # -> Planilla Total (S/.)
        "PROD OPERATIVA",        # -> Planilla (S/.) entre KG Producido
        "VENTA KG X SOL",        # -> Planilla (S/.) entre KG Vendido
        # "Venta Neta (KG)" en el panel aparece como "Suma de Peso total KG" — eso es una
        # COLUMNA con agregación implícita en la UI (Power BI antepone "Suma de" a columnas,
        # no a medidas DAX reales), NO un nombre de medida válido para EVALUATE ROW(). Sin
        # confirmar la tabla/columna real, no se adivina — queda sin este KPI hasta que se
        # confirme con el Analizador de rendimiento (ver nota en build_productividad).
        # Nombres viejos (adivinados, se mantienen como fallback por si alguno coincide):
        "Planilla (S/.) entre KG Vendido", "Planilla Entre KG Vendido", "Planilla/KG Vendido",
        "Planilla (S/.) entre KG Producido", "Planilla Entre KG Producido", "Planilla/KG Producido",
        "Produccion Total (KG)", "Produccion Total KG", "KG Producido", "Produccion Total",
        "Venta Neta (KG)", "Venta Neta KG", "KG Vendido", "Venta Neta",
        "Planilla Total (S/.)", "Planilla Total", "Costo Planilla",
        # Por planta (ATE / PACHACAMAC)
        "Produccion ATE (KG)", "Produccion ATE", "KG Producido ATE",
        "Produccion Pachacamac (KG)", "Produccion Pachacamac", "KG Producido Pachacamac",
        "Venta Neta ATE (KG)", "Venta Neta Pachacamac (KG)",
        "Planilla Total ATE", "Planilla Total Pachacamac",
        "Planilla ATE KG Producido", "Planilla Pachacamac KG Producido",
        # Alternativas
        "Planilla/kg", "S//kg", "Costo Planilla kg",
        "Eficiencia", "HH/kg", "Costo por kg", "S/ por kg",
        "Kg Producidos", "Kg Producción", "Kg Produccion",
        "Planilla Mensual", "Total Planilla", "Gasto Planilla",
        "Ratio Planilla", "Costo Mano de Obra",
        "MOD", "MOD/kg", "Mano de Obra",
    ],
    "planificacion": [
        "Fill Rate", "% Fill Rate", "Tasa Atención", "Tasa Atencion",
        "% Atención", "% Atencion", "Nivel de Servicio", "% Nivel Servicio",
        "Atendido", "% Atendido", "Cumplimiento Pedidos", "% Cumplimiento Pedidos",
        "Pedidos Atendidos", "Pedidos Total", "OC Atendidas",
        "Dead Stock", "% Dead Stock",
        "Avance", "% Avance", "Avance PPTO", "% Avance Presupuesto",
        "Presupuesto", "PPTO", "Ventas PPTO",
        "Ventas Real", "Real vs Ppto", "% Real vs Ppto",
    ],
    "consumo": [
        # Nombres REALES confirmados leyendo el panel Datos > Measures en Power BI Service
        # (Editar informe) — "14. Consumo Materiales indirectos de produccion", 2026-09-03.
        # No validados con DAX Query View (solo lectura de panel) — confirmar antes de usarlos
        # para decisiones críticas.
        "Costo total validado",     # -> Costo Total
        "Peso total KG",            # -> Venta Neta (KG)
        "Ratio costo / kg",         # -> Costo x TN Vendida
        "Producción (KG) Odoo",     # -> Producción Neta (KG)
        "Costo x ton producida",    # -> Costo x TN Producida
        # Medidas del reporte "Consumo de Materiales" (PRODUCCIÓN) — nombres viejos adivinados,
        # se mantienen como fallback:
        "CECO AJUSTADO", "Ceco Ajustado", "CECO", "Costo CECO",
        "MIP", "MIP (S/.)", "MIP Total", "Costo x TN", "Costo por TN",
        "TN PRODUCIDA", "TN Producida", "Toneladas Producidas", "Ton Producida",
        "Consumo Ate", "Consumo Pachacamac", "Consumo Planta",
        "Costo Consumo", "Costo Material", "Costo MP",
        # Medidas ratio consumo/compra
        "Consumo", "Total Consumo", "Importe Consumo", "Monto Consumo",
        "Compras", "Total Compras", "Importe Compras", "Monto Compras",
        "Ratio", "Ratio Consumo", "% Ratio", "Eficiencia", "% Eficiencia",
        "Ratio C/C", "Ratio Compras", "Consumo vs Compras",
        "Cant Compra", "Cant Consumo",
        "Compra Periodo", "Consumo Periodo",
    ],
    "fill_rate": [
        "% FILLRATE", "% Fill Rate", "% FillRate",
        "ORDEN DE VENTA", "FACTURACION", "VENTA PERDIDA",
        "Orden de Venta", "Facturacion", "Venta Perdida",
    ],
}

# ─────────────────────────────────────────────────────────────────────────────

def get_dataset_schema(token, ws_id, dataset_id):
    """Obtiene el schema completo del dataset via REST API (tablas + columnas + medidas)."""
    url = f"{PBI_BASE}/groups/{ws_id}/datasets/{dataset_id}/tables"
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=20)
    if not r.ok:
        return {}
    schema = {}
    for tbl in r.json().get("value", []):
        tbl_name = tbl.get("name", "")
        cols = [c["name"] for c in tbl.get("columns", [])]
        measures = [m["name"] for m in tbl.get("measures", [])]
        schema[tbl_name] = {"columns": cols, "measures": measures}
    return schema

def get_token():
    r = requests.post(TOKEN_URL, data={
        "grant_type": "password", "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET, "username": USERNAME,
        "password": PASSWORD, "scope": PBI_SCOPE,
    }, timeout=30)
    r.raise_for_status()
    return r.json()["access_token"]

def dax(token, ws_id, dataset_id, query, label="query"):
    """Ejecuta una consulta DAX cruda. Retorna lista de filas o []."""
    import time
    url = f"https://api.powerbi.com/v1.0/myorg/groups/{ws_id}/datasets/{dataset_id}/executeQueries"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    body = {"queries": [{"query": query}], "serializerSettings": {"includeNulls": True}}
    for attempt in range(3):
        try:
            r = requests.post(url, json=body, headers=headers, timeout=30)
            if r.status_code == 429:
                wait = int(r.headers.get("Retry-After", 10))
                print(f"    [dax:{label}] throttled — esperando {wait}s...")
                time.sleep(wait)
                continue
            if r.status_code != 200:
                print(f"    [dax:{label}] {r.status_code} {r.text[:120]}")
                return []
            tables = r.json().get("results", [{}])[0].get("tables", [])
            return tables[0].get("rows", []) if tables else []
        except Exception as e:
            print(f"    [dax:{label}] error: {e}")
            if attempt < 2:
                time.sleep(3)
    return []

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
    """Escanea medidas reales. Aplica filtro de mes anterior si hay tabla de fecha conocida."""
    candidates = SCAN_CANDIDATES.get(ds_key, [])
    date_ctx = DATE_CONTEXT.get(ds_key)  # (tabla, columna) o None
    found = {}

    def get_val(name):
        if date_ctx:
            v = dax_prev_month(token, ws_id, dataset_id, name, date_ctx[0], date_ctx[1], f"{ds_key}_{name[:15]}")
            if v is not None:
                return v, True
        return try_measure(token, ws_id, dataset_id, name)

    cached_found = (cache or {}).get(ds_key, {})
    if cached_found:
        print(f"    Verificando {len(cached_found)} medidas en caché...")
        for name in list(cached_found.keys()):
            v, exists = get_val(name)
            if exists:
                found[name] = v
                print(f"      ✓ [{name}] = {v} (caché)")
        scanned_before = set(cached_found.keys())
        candidates = [c for c in candidates if c not in scanned_before]
    print(f"    Escaneando {len(candidates)} candidatos nuevos...")
    for name in candidates:
        v, exists = get_val(name)
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
    # Nombres confirmados: "Cuentas x Pagar", "Refinanciamiento", "Proveedores"
    total_val = (found.get("Cuentas x Pagar") or found.get("CUENTAS X PAGAR") or
                 found.get("CxP Total") or found.get("Total CxP") or found.get("Saldo CxP") or
                 found.get("Deuda Total"))
    refin_val = (found.get("Refinanciamiento") or found.get("Refinanciado") or
                 found.get("Saldo Refinanciado"))
    proveed_val = found.get("Proveedores") or found.get("Nro Proveedores")
    # Tramos vencido
    vigente_val = found.get("Vigente CxP") or found.get("CxP Vigente") or found.get("Saldo Vigente")
    venc15_val  = found.get("Vencido < 15") or found.get("Vencido Menor 15d") or found.get("Vencido 0-15d")
    venc30_val  = found.get("Vencido 16-30d") or found.get("Vencido 30d") or found.get("Vencido 1-30")
    venc90_val  = found.get("Vencido 31-90") or found.get("Vencido 31-90d")
    # Días CxP
    dias_val = (found.get("Dias CxP") or found.get("DPP") or found.get("Dias Pago") or
                found.get("Rotacion CxP") or found.get("Dias Promedio Pago") or
                found.get("Dias de Pago") or found.get("Días de Pago") or
                found.get("PPP") or found.get("Plazo Pago"))

    total = to_float(total_val)
    refin = to_float(refin_val)
    dias  = to_float(dias_val)

    # Semáforo: si refinanciado > 30% del total, amarillo; dias > 90, rojo
    sem = "green"
    if dias and dias > 90: sem = "red"
    elif dias and dias > 60: sem = "yellow"
    elif refin and total and refin / total > 0.30: sem = "yellow"

    kpis = []
    if total_val is not None:
        kpis.append({"label": "CxP Total", "valor": fmt_soles(total_val)})
    if vigente_val is not None:
        kpis.append({"label": "Vigente", "valor": fmt_soles(vigente_val), "estado": "green"})
    if venc15_val is not None:
        kpis.append({"label": "Vencido <15d", "valor": fmt_soles(venc15_val), "estado": "yellow"})
    if venc30_val is not None:
        kpis.append({"label": "Vencido 16-30d", "valor": fmt_soles(venc30_val), "estado": "red"})
    if venc90_val is not None:
        kpis.append({"label": "Vencido 31-90d", "valor": fmt_soles(venc90_val), "estado": "red"})
    if refin_val is not None:
        s = "yellow" if refin and total and refin/total > 0.30 else "green"
        kpis.append({"label": "Refinanciado", "valor": fmt_soles(refin_val), "estado": s})
    if proveed_val is not None:
        kpis.append({"label": "# Proveedores", "valor": str(int(to_float(proveed_val) or 0))})
    if dias is not None:
        s_d = "red" if dias > 90 else ("yellow" if dias > 60 else "green")
        kpis.append({"label": "Días CxP", "valor": f"{dias:.0f}d", "meta": "<90d", "estado": s_d})

    alerta = f"CxP {dias:.0f}d — revisar flujo" if dias and dias > 60 else (
             f"Refinanciado {fmt_soles(refin_val)} — gestionar" if refin and total and refin/total > 0.30 else None)
    return {"estado": sem, "alerta": alerta, "kpis": kpis}, dias

def build_margen(found):
    margen_val = (found.get("% Margen") or found.get("R. MARGEN") or found.get("R. Margen") or
                  found.get("Margen Variable") or found.get("% Margen Variable") or
                  found.get("MV") or found.get("% MV") or found.get("MV%") or
                  found.get("Margen Variable %") or found.get("Margen") or
                  found.get("% MV PAUNO") or found.get("Pct MV"))
    costo_total_val = found.get("Costo Total")
    ventas_val = (found.get("Ventas Mes Actual") or found.get("Ventas Actuales") or
                  found.get("Ventas") or found.get("Venta Total") or found.get("Venta Neta"))
    precio_val = (found.get("Precio/kg") or found.get("Precio kg") or found.get("Precio Promedio kg") or
                  found.get("Precio x Kilo") or found.get("Precio Kilo"))
    costo_val  = (found.get("Costo/kg") or found.get("Costo kg") or
                  found.get("R. COSTO UNIT") or found.get("Costo x Kilo") or found.get("Costo Kilo"))

    margen_pct = to_float(margen_val)
    if margen_pct and abs(margen_pct) < 1: margen_pct *= 100
    # Si no hay medida directa de margen, calcular desde Precio/kg y Costo/kg
    if margen_pct is None and precio_val is not None and costo_val is not None:
        p, c = to_float(precio_val), to_float(costo_val)
        if p and c and p > 0:
            margen_pct = (p - c) / p * 100
    sem = sem_thresh(margen_pct, red_below=46, yellow_below=52)

    kpis = []
    if ventas_val is not None:
        kpis.append({"label": "Ventas mes", "valor": fmt_soles(ventas_val)})
    if margen_pct is not None:
        kpis.append({"label": "Margen variable", "valor": f"{margen_pct:.1f}%", "meta": "52%", "estado": sem})
    if precio_val is not None:
        kpis.append({"label": "Precio/kg", "valor": f"S/{to_float(precio_val):.2f}"})
    if costo_val is not None:
        kpis.append({"label": "Costo/kg", "valor": f"S/{to_float(costo_val):.2f}"})
    if costo_total_val is not None:
        kpis.append({"label": "Costo Total", "valor": fmt_soles(costo_total_val)})

    alerta = f"Margen {fmt_pct(margen_pct)} — bajo meta 52%" if sem != "green" and margen_pct else None
    return {"estado": sem, "alerta": alerta, "kpis": kpis}, ventas_val, margen_pct

def build_mermas(found):
    ate_val  = found.get("% Merma Ate") or found.get("Merma Ate") or found.get("Merma Planta Ate")
    pach_val = found.get("% Merma Pachacamac") or found.get("Merma Pachacamac") or found.get("Merma Lurin")
    tot_val  = (found.get("% Merma Total") or found.get("% Merma") or
                found.get("% MERMAS") or found.get("Merma %") or
                found.get("Tasa Merma") or found.get("Merma Total"))
    terc_val = found.get("% Merma Terceros") or found.get("Merma Terceros")
    # UEN (Control de Producción: B&D, TIGO, MAQUILA)
    bd_val   = found.get("% Merma B&D") or found.get("Merma B&D") or found.get("% Merma BD")
    tigo_val = found.get("% Merma Tigo") or found.get("Merma Tigo") or found.get("% Merma TIGO")
    maq_val  = found.get("% Merma Maquila") or found.get("Merma Maquila") or found.get("% Merma MAQUILA")
    # Detalle: REAL, STD, DESVIACION (tabla "Detalle de Mermas por Planta")
    real_val  = found.get("REAL")
    std_val   = found.get("STD")
    desv_val  = found.get("DESVIACION")
    item_val  = found.get("ITEM DESVIADO")

    def pct(v):
        n = to_float(v)
        if n is None: return None
        if abs(n) < 1: n *= 100
        return n

    ate, pach, tot, terc = pct(ate_val), pct(pach_val), pct(tot_val), pct(terc_val)
    bd, tigo, maq = pct(bd_val), pct(tigo_val), pct(maq_val)
    worst = max(filter(lambda x: x is not None, [ate, pach, tot, terc]), default=None)
    sem = "green"
    if worst:
        if worst > 3: sem = "red"
        elif worst > 2: sem = "yellow"

    kpis = []
    for label, val, meta in [
        ("Merma Total",       tot,  "2%"),
        ("Merma Ate",         ate,  "2%"),
        ("Merma Pachacamac",  pach, "2%"),
        ("Merma Terceros",    terc, "2%"),
        ("Merma B&D",         bd,   "2%"),
        ("Merma TIGO",        tigo, "2%"),
        ("Merma MAQUILA",     maq,  "2%"),
    ]:
        if val is not None:
            s = "red" if val > 3 else ("yellow" if val > 2 else "green")
            kpis.append({"label": label, "valor": f"{val:.2f}%", "meta": meta, "estado": s})

    if real_val  is not None: kpis.append({"label": "REAL (unidades)", "valor": str(real_val)})
    if std_val   is not None: kpis.append({"label": "STD (estándar)", "valor": str(std_val)})
    if desv_val  is not None: kpis.append({"label": "Desviación", "valor": str(desv_val)})
    if item_val  is not None: kpis.append({"label": "Ítems Desviados", "valor": str(item_val)})

    alerta = f"Merma {worst:.2f}% — sobre meta 2%" if sem != "green" and worst else None
    # Guardar por_uen para el frontend
    por_uen = []
    for uen, val in [("B&D", bd), ("TIGO", tigo), ("MAQUILA", maq)]:
        if val is not None:
            s = "red" if val > 3 else ("yellow" if val > 2 else "green")
            por_uen.append({"uen": uen, "merma": f"{val:.2f}%", "estado": s})
    # Guardar por_planta
    por_planta = []
    for planta, val in [("ATE", ate), ("PACHACAMAC", pach), ("TERCEROS", terc)]:
        if val is not None:
            s = "red" if val > 3 else ("yellow" if val > 2 else "green")
            por_planta.append({"planta": planta, "merma": f"{val:.2f}%", "estado": s})

    result = {"estado": sem, "alerta": alerta, "kpis": kpis}
    if por_uen:    result["por_uen"]    = por_uen
    if por_planta: result["por_planta"] = por_planta
    return result

def build_consumo_materiales(found):
    """Reporte '14. Consumo Materiales indirectos de produccion' — PAUNO.

    Estado por KPI:
      Costo Total          -> VALIDADO (dax_consumo_costo_total). Filtrado por
                              Calendario[Año]=PREV_YEAR desde 2026-09-04 — la primera
                              captura (2026-09-03) no tenía año seleccionado en el
                              reporte y traía el acumulado histórico total; se repitió
                              la captura con "2026" marcado y se corrigió.
      Venta Neta KG        -> VALIDADO (dax_consumo_venta_neta_kg). Filtro de fecha
                              propio (desde 2025-07-01).
      Costo x TN Vendida   -> VALIDADO (dax_consumo_costo_x_tn_vendida). Mismos
                              7 filtros que Venta Neta KG.
      Producción Neta KG   -> VALIDADO (dax_consumo_produccion_neta_kg). Mismos
                              filtros que Costo Total, incluido Calendario[Año] —
                              confirmado 1:1 con su propia Copiar consulta (2026-09-04),
                              idéntica a la de Costo Total salvo la medida.
      Costo x TN Producida -> VALIDADO (dax_consumo_costo_x_tn_producida). Mismos
                              7 filtros que Venta Neta KG / Costo x TN Vendida.

    Los 5 KPIs de este reporte quedaron validados 1:1 con Copiar consulta.
    """
    costo_total_val = found.get("Costo total validado") or found.get("Costo Consumo")
    venta_kg_val    = found.get("Peso total KG") or found.get("Venta Neta (KG)") or found.get("Venta Neta KG")
    costo_x_tn_vend_val = found.get("Ratio costo / kg") or found.get("Costo x TN") or found.get("Costo por TN")
    prod_neta_val   = found.get("Producción (KG) Odoo") or found.get("TN PRODUCIDA") or found.get("TN Producida")
    costo_x_tn_prod_val = found.get("Costo x ton producida")

    costo_total     = to_float(costo_total_val)
    venta_kg        = to_float(venta_kg_val)
    costo_x_tn_vend = to_float(costo_x_tn_vend_val)
    prod_neta       = to_float(prod_neta_val)
    costo_x_tn_prod = to_float(costo_x_tn_prod_val)

    # Sin meta ni benchmark definido por negocio todavía — verde de referencia.
    sem = "green"

    kpis = []
    if costo_total is not None:
        kpis.append({"label": "Costo Total", "valor": fmt_soles(abs(costo_total)),
                     "validado": True, "nota": "Filtro exacto confirmado — sin corte de fecha (histórico total)"})
    if venta_kg is not None:
        kpis.append({"label": "Venta Neta (KG)", "valor": f"{venta_kg:,.0f} KG", "validado": True})
    if costo_x_tn_vend is not None:
        kpis.append({"label": "Costo x TN Vendida", "valor": f"S/{costo_x_tn_vend:.2f}", "validado": True})
    if prod_neta is not None:
        kpis.append({"label": "Producción Neta (KG)", "valor": f"{prod_neta:,.0f} KG", "validado": True})
    if costo_x_tn_prod is not None:
        kpis.append({"label": "Costo x TN Producida", "valor": f"S/{costo_x_tn_prod:.2f}", "validado": True})

    if not kpis:
        return None  # No hay datos — no agregar el reporte
    # Marcador de reporte: True mientras quede al menos un KPI sin validar con
    # Copiar consulta. Cada KPI individual lleva su propio campo "validado".
    hay_sin_validar = any(not k.get("validado") for k in kpis)
    return {"estado": sem, "alerta": None, "kpis": kpis, "no_validado_dax": hay_sin_validar}

def build_compras(found):
    ratio_val  = (found.get("Ratio") or found.get("Eficiencia") or found.get("Eficiencia Costo") or
                  found.get("Ratio C/V") or found.get("Ratio Compras") or
                  found.get("Ratio Consumo/Compra") or found.get("% Ratio") or found.get("Consumo/Compra"))
    consumo_val = (found.get("Cant Consumo") or found.get("Consumo") or
                   found.get("Total Consumo") or found.get("Importe Consumo"))
    compra_val  = (found.get("Cant Compra") or found.get("Valor Compras") or
                   found.get("Compras") or found.get("Total Compras") or found.get("Importe Compras"))
    # Medidas del Reporte Compras — ANALISIS DE MATERIALES (PLANIFICACION Y COMPRA)
    stock_pp_val    = found.get("Stock PP")
    stock_val_val   = found.get("Stock Valorizado")
    cons_3m_val     = found.get("Consumo Prom 3M")
    cons_6m_val     = found.get("Consumo Prom 6M")
    faltantes_val   = found.get("Faltantes")
    lead_time_val   = found.get("Lead Time")

    ratio = to_float(ratio_val)
    if ratio and abs(ratio) < 2: ratio *= 100
    if ratio is None and consumo_val and compra_val:
        c, p = to_float(consumo_val), abs(to_float(compra_val) or 0)
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
    # Métricas de planificación de compras
    if stock_pp_val is not None:
        kpis.append({"label": "Stock PP (Punto de Pedido)", "valor": str(int(to_float(stock_pp_val) or 0))})
    if stock_val_val is not None:
        kpis.append({"label": "Stock Valorizado", "valor": fmt_soles(stock_val_val)})
    if cons_3m_val is not None:
        kpis.append({"label": "Consumo Prom 3M", "valor": str(int(to_float(cons_3m_val) or 0))})
    if cons_6m_val is not None:
        kpis.append({"label": "Consumo Prom 6M", "valor": str(int(to_float(cons_6m_val) or 0))})
    if faltantes_val is not None:
        n = to_float(faltantes_val)
        s = "red" if (n or 0) > 0 else "green"
        kpis.append({"label": "Faltantes", "valor": str(int(n or 0)), "estado": s})
    if lead_time_val is not None:
        kpis.append({"label": "Lead Time Prom (días)", "valor": f"{to_float(lead_time_val):.0f}d"})

    # Semáforo: si hay faltantes o sin ratio, usar faltantes
    if not ratio and faltantes_val is not None:
        n = to_float(faltantes_val)
        sem = "red" if (n or 0) > 10 else ("yellow" if (n or 0) > 0 else "green")

    alerta = f"Ratio {ratio:.1f}% — {interp}" if sem != "green" and ratio else (
             f"Faltantes: {int(to_float(faltantes_val) or 0)} ítems" if faltantes_val and to_float(faltantes_val) else None)
    return {"estado": sem, "alerta": alerta, "kpis": kpis}, ratio

def build_inventario(found):
    # ── Composición confirmada con Copiar consulta (2026-09-04): lista de
    # (clasificación, saldo soles) desde 'SALDO ACTUAL'[Clasificación Segun
    # Consumo] — WORKING / EXCESO 1 / EXCESO 2 / DEAD. Tiene prioridad sobre
    # los nombres adivinados de abajo, que quedan solo como fallback.
    composicion = found.get("_composicion") or []
    comp_map = {}
    for clasif, saldo in composicion:
        key = str(clasif or "").strip().upper()
        comp_map[key] = to_float(saldo)

    dead_val    = comp_map.get("DEAD") or found.get("Dead Stock") or found.get("Stock Muerto") or found.get("Inmovilizado") or found.get("Stock Inmovilizado")
    working_val = comp_map.get("WORKING") or found.get("Working Stock") or found.get("Stock Activo") or found.get("Stock Working") or found.get("Stock Normal")
    exceso1_val = comp_map.get("EXCESO 1")
    exceso2_val = comp_map.get("EXCESO 2")
    total_val   = (sum(v for v in comp_map.values() if v) if comp_map else None) or (
                   found.get("Inventario Total") or found.get("Total Inventario") or
                   found.get("Saldo Inventario") or found.get("Valor Inventario") or
                   found.get("Stock Total") or found.get("Total Stock") or found.get("Costo Inventario"))
    cob_total = found.get("Cobertura Total (días)")
    cob_mp    = found.get("Cobertura MP (días)")

    deadpct_val = found.get("% Dead Stock") or found.get("% Inmovilizado") or found.get("% Dead") or found.get("Pct Dead")
    dead_pct = to_float(deadpct_val)
    if dead_pct and abs(dead_pct) < 1: dead_pct *= 100
    if dead_pct is None and dead_val and total_val:
        d, t = to_float(dead_val), to_float(total_val)
        if d and t and t > 0: dead_pct = d / t * 100

    # No usar sem_thresh() aquí: su heurística "abs(n)<1 -> multiplicar x100"
    # duplicaría la escala cuando dead_pct ya es un porcentaje válido y chico
    # (ej. 0.3%), convirtiéndolo por error en 30% y disparando rojo.
    if dead_pct is None:
        sem = "green"
    elif dead_pct >= 10:
        sem = "red"
    elif dead_pct >= 5:
        sem = "yellow"
    else:
        sem = "green"

    kpis = []
    if total_val is not None:
        kpis.append({"label": "Inventario Total", "valor": fmt_soles(total_val)})
    if cob_total is not None:
        kpis.append({"label": "Cobertura Total", "valor": f"{cob_total:.0f} días"})
    if cob_mp is not None:
        kpis.append({"label": "Cobertura MP", "valor": f"{cob_mp:.0f} días"})
    if dead_val is not None:
        kpis.append({"label": "Dead Stock", "valor": fmt_soles(dead_val), "meta": "<S/500K", "estado": sem})
    if dead_pct is not None:
        kpis.append({"label": "% Dead Stock", "valor": f"{dead_pct:.1f}%", "meta": "<5%", "estado": sem})
    if working_val is not None:
        kpis.append({"label": "Working Stock", "valor": fmt_soles(working_val)})
    if exceso1_val is not None:
        kpis.append({"label": "Exceso 1 (2-5 meses)", "valor": fmt_soles(exceso1_val)})
    if exceso2_val is not None:
        kpis.append({"label": "Exceso 2 (5-12 meses)", "valor": fmt_soles(exceso2_val)})

    alerta = f"Dead Stock {dead_pct:.1f}% del inventario" if sem != "green" and dead_pct else None
    return {"estado": sem, "alerta": alerta, "kpis": kpis}

def build_control(found):
    """Dashboard de Control Interno ('8. Reporte de auditoría').

    Confirmado con Copiar consulta el 2026-09-04, filtro único Empresa="Pauno",
    sin fecha (acumulado histórico total): Satisfactorio, Con Observaciones,
    Critico — las 3 comparten el mismo filtro y patrón de medida.

    "% Cumplimiento" (69.32% en pantalla = Calificación 653 ÷ Puntos Teóricos
    942) queda pendiente: son 2 medidas más que no se pudieron capturar con
    Copiar consulta en esta sesión — no se inventa el cálculo con los 3
    conteos de aquí porque no es la misma fórmula (no es una simple división
    de conteos, es una calificación ponderada).
    """
    satisf_val = found.get("Satisfactorio")
    obs_val    = found.get("Con Observaciones")
    crit_val   = found.get("Critico")
    abiertos_val = found.get("Planes Abiertos")
    cerrados_val = found.get("Planes Cerrados")
    ejecutado_val = found.get("Ejecutado")
    total_planes_val = found.get("Total Planes de Acción")
    atrasados_val = found.get("Planes Atrasados")

    satisf = to_float(satisf_val)
    obs    = to_float(obs_val)
    crit   = to_float(crit_val)
    abiertos = to_float(abiertos_val)
    cerrados = to_float(cerrados_val)
    ejecutado = to_float(ejecutado_val)
    total  = sum(v for v in (satisf, obs, crit) if v is not None) or None
    # Total real del modelo (medida propia) en vez de sumar Abiertos+Cerrados a mano.
    total_planes = to_float(total_planes_val) or ((abiertos or 0) + (cerrados or 0) if (abiertos is not None or cerrados is not None) else None)
    atrasados = to_float(atrasados_val)
    avance_planes_pct = (cerrados / total_planes * 100) if (cerrados is not None and total_planes) else None

    crit_pct = (crit / total * 100) if (crit is not None and total) else None
    sem = "green"
    if crit_pct is not None:
        sem = "red" if crit_pct > 20 else ("yellow" if crit_pct > 10 else "green")

    kpis = []
    if total is not None:
        kpis.append({"label": "Puntos de Control evaluados", "valor": str(int(total))})
    if satisf is not None:
        kpis.append({"label": "Satisfactorio", "valor": str(int(satisf)), "estado": "green"})
    if obs is not None:
        kpis.append({"label": "Con Observaciones", "valor": str(int(obs)), "estado": "yellow"})
    if crit is not None:
        kpis.append({"label": "Crítico", "valor": str(int(crit)), "meta": "0", "estado": "red" if crit > 0 else "green"})
    if crit_pct is not None:
        kpis.append({"label": "% Puntos en estado Crítico", "valor": f"{crit_pct:.1f}%", "meta": "<10%", "estado": sem})
    if total_planes is not None:
        kpis.append({"label": "Total Planes de Acción", "valor": str(int(total_planes))})
    if abiertos is not None:
        kpis.append({"label": "Planes de Acción Abiertos", "valor": str(int(abiertos)), "estado": "yellow" if abiertos > 0 else "green"})
    if cerrados is not None:
        kpis.append({"label": "Planes de Acción Cerrados", "valor": str(int(cerrados)), "estado": "green"})
    if atrasados is not None:
        kpis.append({"label": "Planes de Acción Atrasados", "valor": str(int(atrasados)), "meta": "0", "estado": "red" if atrasados > 0 else "green"})
    if avance_planes_pct is not None:
        kpis.append({"label": "% Avance Planes de Acción", "valor": f"{avance_planes_pct:.1f}%", "estado": "green" if avance_planes_pct>=80 else ("yellow" if avance_planes_pct>=50 else "red")})
    if ejecutado is not None:
        kpis.append({"label": "Puntos Ejecutados", "valor": str(int(ejecutado))})

    alerta = f"{int(crit)} puntos de control en estado Crítico ({crit_pct:.1f}% del total)" if sem != "green" and crit else None
    return {"estado": sem if kpis else "green", "alerta": alerta, "kpis": kpis}

def build_productividad(found):
    """Reporte '13. Productividad por Funcionario' — PAUNO.

    Nombres de medida confirmados el 2026-09-03 leyendo el panel Datos > Measures
    en Power BI Service (Editar informe) y validados con EVALUATE ROW(...) en la
    Vista de consultas DAX (todas resuelven sin error):
      SUM PROD          -> Producción Total (KG)
      GASTO TOTAL        -> Planilla Total (S/.)
      PROD OPERATIVA      -> Planilla (S/.) entre KG Producido
      VENTA KG X SOL       -> Planilla (S/.) entre KG Vendido

    "Venta Neta (KG)" aparece en el panel como "Suma de Peso total KG" porque es la
    COLUMNA 'Maestra de Facturacion (Total)'[Peso total K] con SUM directo, no una
    medida DAX — confirmado el 2026-09-03 con Analizador de rendimiento > Copiar
    consulta (ver dax_productividad_venta_neta_kg). Trae 2 filtros propios que las
    otras 4 tarjetas no tienen: CATEGORIZACION="VENTA BRUTA" y TIPO DE NEGOCIO N2
    no vacío.
    """
    # ── BARRERA DE SEGURIDAD — historial ────────────────────────────────────
    # GASTO TOTAL se probó en vivo (2026-09-03) sin el filtro real del reporte:
    #   sin filtro     = -247
    #   con solo Calendario[Año]=2026 = -117
    #   tarjeta real   = S/4,484,364
    # Confirmado con "Copiar consulta" (Analizador de rendimiento) que el reporte
    # aplica 9 filtros: Calendario[Año]=2026, Calendario[MesActual]="Otros",
    # Exl Cuenta Contables[CLASIFICACION] IN {"AMBAS","GASTO DE PERSONAL OPERATIVO"},
    # [GASTO_PERSONAL]="GASTO DE PERSONAL", [SUB_CATEGORIA]="Gasto de Personal Operativo",
    # exclusión de categoria_producto/producto/estado, y Calendario[Date] > 2025-07-31.
    # main() ahora llama dax_productividad_pauno() con esos 9 filtros exactos y
    # sobrescribe found[...] con el valor correcto — por eso ya es seguro leer
    # found.get(...) directo aquí. GASTO_TOTAL_CONFIRMADO queda como bandera
    # histórica del hallazgo, ya no bloquea nada.
    GASTO_TOTAL_CONFIRMADO = True

    plan_kg_vend = plan_kg_prod = plan_total = None
    if GASTO_TOTAL_CONFIRMADO:
        plan_kg_vend  = (found.get("VENTA KG X SOL") or
                         found.get("Planilla (S/.) entre KG Vendido") or
                         found.get("Planilla Entre KG Vendido") or found.get("Planilla/KG Vendido"))
        plan_kg_prod  = (found.get("PROD OPERATIVA") or
                         found.get("Planilla (S/.) entre KG Producido") or
                         found.get("Planilla Entre KG Producido") or found.get("Planilla/KG Producido"))
        plan_total    = (found.get("GASTO TOTAL") or
                         found.get("Planilla Total (S/.)") or found.get("Planilla Total"))

    prod_total    = (found.get("SUM PROD") or
                     found.get("Produccion Total (KG)") or found.get("Produccion Total KG") or
                     found.get("Produccion Total") or found.get("Kg Producidos") or found.get("Kg Producción"))
    # "Venta Neta (KG)" — columna SUM confirmada y filtrada por
    # dax_productividad_venta_neta_kg(), guardada bajo esta clave interna.
    venta_neta_kg = found.get("__VENTA_NETA_KG__")

    kpis = []
    def sf2(v, prefix="S/"):
        n = to_float(v)
        return f"{prefix}{n:.2f}" if n else None
    def kg_fmt(v):
        n = to_float(v)
        return f"{n:,.0f} kg" if n else None

    if plan_kg_prod is not None:
        kpis.append({"label": "Planilla S/. / KG Producido", "valor": sf2(plan_kg_prod) or "—"})
    if plan_kg_vend is not None:
        kpis.append({"label": "Planilla S/. / KG Vendido",   "valor": sf2(plan_kg_vend) or "—"})
    if prod_total is not None:
        kpis.append({"label": "Producción Total (KG)",        "valor": kg_fmt(prod_total) or "—"})
    if venta_neta_kg is not None:
        kpis.append({"label": "Venta Neta (KG)",              "valor": kg_fmt(venta_neta_kg) or "—"})
    if plan_total is not None:
        kpis.append({"label": "Planilla Total",               "valor": fmt_soles(plan_total)})

    # Tendencia: menor S//kg = mayor eficiencia
    n_prod = to_float(plan_kg_prod)
    sem = "green"
    if n_prod:
        if n_prod > 0.60: sem = "red"
        elif n_prod > 0.50: sem = "yellow"
    alerta = f"Planilla/KG Producido S/{n_prod:.2f} — revisar eficiencia" if sem != "green" and n_prod else None
    return {"estado": sem, "alerta": alerta, "kpis": kpis}

def build_fill_rate(found):
    # '% Fill Rate' SÍ responde al filtro de mes; '% FILLRATE' devuelve el acumulado
    # histórico igual en todos los meses. Se prefiere la que refleja el mes en curso.
    fill_val = (found.get("% Fill Rate") or found.get("% FILLRATE") or found.get("% FillRate") or
                found.get("Fill Rate") or found.get("Tasa Atención") or found.get("Tasa Atencion"))
    fill_pct = to_float(fill_val)
    if fill_pct and abs(fill_pct) < 1: fill_pct *= 100
    sem = sem_thresh(fill_pct, red_below=85, yellow_below=95)

    ov_val     = found.get("ORDEN DE VENTA") or found.get("Orden de Venta")
    fact_val   = found.get("FACTURACION") or found.get("Facturacion")
    vp_val     = found.get("VENTA PERDIDA") or found.get("Venta Perdida")
    ov         = to_float(ov_val)
    fact       = to_float(fact_val)
    vp         = to_float(vp_val)

    kpis = []
    if fill_pct is not None:
        kpis.append({"label": "Fill Rate", "valor": f"{fill_pct:.1f}%", "meta": "98%", "estado": sem})
    if ov:
        kpis.append({"label": "Orden de Venta", "valor": fmt_soles(ov)})
    if fact:
        kpis.append({"label": "Facturación", "valor": fmt_soles(fact)})
    if vp:
        kpis.append({"label": "Venta Perdida", "valor": fmt_soles(vp), "estado": "red" if vp > 0 else "green"})

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

        # Cargar medidas descubiertas (fuente de verdad del discover workflow)
        discovered_path = OUTPUT_DIR / "discovered_measures.json"
        discovered = {}
        if discovered_path.exists():
            try:
                discovered = json.loads(discovered_path.read_text()).get("datasets", {})
                print(f"  Usando discovered_measures.json ({len(discovered)} datasets)")
            except: pass

        # Escanear medidas reales de cada dataset
        measure_cache = load_measure_cache()
        scanned = {}
        for ds_key, did in ids.items():
            print(f"  Escaneando {ds_key}...")
            scanned[ds_key] = scan_dataset(token, ws_id, did, ds_key, measure_cache)
            # Completar con valores del discover (si el scan no encontró algo)
            disc_ds = discovered.get(ds_key, {})
            for m_name, m_vals in disc_ds.items():
                if m_name not in scanned.get(ds_key, {}):
                    val = m_vals.get("dated") if m_vals.get("dated") is not None else m_vals.get("plain")
                    if val is not None:
                        scanned.setdefault(ds_key, {})[m_name] = val
                        print(f"    [{m_name}] = {val} (discovered)")

        # ── Compras: filtrar via TREATAS sobre tabla Calendario (correcto para Live Connection)
        if "compras" in ids:
            ds_c = ids["compras"]
            print(f"  Compras: buscando valor agosto {PREV_YEAR}-{PREV_MONTH:02d} via TREATAS...")
            compras_val = None
            # TREATAS inyecta un rango de fechas como filtro via la relación con Calendario
            # Probamos las variantes de nombre de columna en la tabla Calendario
            for cal_col in ["Date", "Fecha", "fecha", "date", "CalendarDate", "DateKey"]:
                for cal_tbl in ["Calendario", "Calendar", "dCalendario", "Fechas", "Dim_Fecha"]:
                    q = f"""EVALUATE ROW("v",
  CALCULATE(
    SUM('Compras'[data.monto_total_linea]),
    TREATAS(
      FILTER(
        GENERATESERIES(DATE({PREV_YEAR},{PREV_MONTH},1), DATE({PREV_YEAR},{PREV_MONTH},31), 1),
        [Value] <= DATE({PREV_YEAR},{PREV_MONTH},31)
      ),
      '{cal_tbl}'[{cal_col}]
    )
  ))"""
                    rows = dax(token, ws_id, ds_c, q, f"treatas_{cal_tbl[:6]}_{cal_col[:4]}")
                    if rows:
                        v = rows[0].get("[v]") or rows[0].get("v")
                        if v is not None:
                            fv = float(v or 0)
                            print(f"    TREATAS '{cal_tbl}'[{cal_col}] → {v}")
                            if 0 < abs(fv) < 50_000_000:
                                compras_val = v
                                print(f"    ✓ Compras agosto: {v}")
                                break
                    if compras_val is not None:
                        break
                if compras_val is not None:
                    break
            if compras_val is not None:
                scanned.setdefault("compras", {})["Valor Compras"] = compras_val
            # No usar fallback total — el SUM acumulado es incorrecto para mostrar como "Compras mes"
            if compras_val is None:
                print(f"    Compras: ratio agosto no accesible via API (Live Connection)")

        # ── Consumo: 'Costo Total' con el filtro EXACTO confirmado (Copiar consulta, 2026-09-03)
        # Sin filtro de fecha — ver docstring de dax_consumo_costo_total. Sobrescribe
        # cualquier valor sin filtrar que traiga el heurístico de abajo.
        if "consumo" in ids:
            v_ct = dax_consumo_costo_total(token, ws_id, ids["consumo"])
            if v_ct is not None:
                scanned.setdefault("consumo", {})["Costo total validado"] = v_ct
                print(f"    ✓ Consumo [Costo total validado] filtrado = {v_ct}")
            else:
                print("    ✗ Consumo [Costo total validado] — la consulta filtrada no devolvió valor")

            v_vn = dax_consumo_venta_neta_kg(token, ws_id, ids["consumo"])
            if v_vn is not None:
                scanned.setdefault("consumo", {})["Peso total KG"] = v_vn
                print(f"    ✓ Consumo [Venta Neta KG] filtrado = {v_vn}")
            else:
                print("    ✗ Consumo [Venta Neta KG] — la consulta filtrada no devolvió valor")

            v_ctv = dax_consumo_costo_x_tn_vendida(token, ws_id, ids["consumo"])
            if v_ctv is not None:
                scanned.setdefault("consumo", {})["Ratio costo / kg"] = v_ctv
                print(f"    ✓ Consumo [Costo x TN Vendida] filtrado = {v_ctv}")
            else:
                print("    ✗ Consumo [Costo x TN Vendida] — la consulta filtrada no devolvió valor")

            v_pn = dax_consumo_produccion_neta_kg(token, ws_id, ids["consumo"])
            if v_pn is not None:
                scanned.setdefault("consumo", {})["Producción (KG) Odoo"] = v_pn
                print(f"    ✓ Consumo [Producción Neta KG] filtrado = {v_pn}")
            else:
                print("    ✗ Consumo [Producción Neta KG] — la consulta filtrada no devolvió valor")

            v_ctp = dax_consumo_costo_x_tn_producida(token, ws_id, ids["consumo"])
            if v_ctp is not None:
                scanned.setdefault("consumo", {})["Costo x ton producida"] = v_ctp
                print(f"    ✓ Consumo [Costo x TN Producida] filtrado = {v_ctp}")
            else:
                print("    ✗ Consumo [Costo x TN Producida] — la consulta filtrada no devolvió valor")

        # ── Consumo: SUM directo desde tablas del dataset consumo
        if "consumo" in ids:
            print("  Consumo: descubriendo tablas...")
            consumo_tables = discover_tables_in_dataset(token, ws_id, ids["consumo"])
            if consumo_tables:
                print(f"    Tablas consumo: {list(consumo_tables.keys())}")
            for tbl_name in consumo_tables:
                cols_c = consumo_tables[tbl_name]
                # Buscar columna de monto/importe/costo
                for col in cols_c:
                    col_lower = col.lower()
                    if any(kw in col_lower for kw in ["monto", "importe", "costo", "valor", "consumo", "total"]):
                        col_bare = col.split("[")[-1].rstrip("]")
                        rows_c = dax(token, ws_id, ids["consumo"],
                                     f"EVALUATE ROW(\"v\", SUM('{tbl_name}'[{col_bare}]))",
                                     f"consumo_sum_{col_bare[:20]}")
                        if rows_c:
                            cv = rows_c[0].get("[v]") or rows_c[0].get("v")
                            if cv is not None and float(cv or 0) != 0:
                                scanned.setdefault("consumo", {})["Total Consumo"] = cv
                                print(f"    Consumo SUM '{col}': {cv}")
                                break
                if scanned.get("consumo", {}).get("Total Consumo"):
                    break

        # ── Productividad: filtro EXACTO confirmado con "Copiar consulta" (2026-09-03)
        # Sobrescribe cualquier valor sin filtrar que haya traído el scan genérico —
        # ese valor sin filtro es garbage (ver dax_productividad_pauno). Solo se piden
        # las 4 medidas confirmadas y validadas con esta consulta exacta.
        if "productividad_ds" in ids:
            print("  Productividad: aplicando filtro exacto confirmado (Copiar consulta)...")
            for m_name, label in [
                ("SUM PROD", "prod_sumprod"),
                ("GASTO TOTAL", "prod_gastototal"),
                ("PROD OPERATIVA", "prod_prodoperativa"),
                ("VENTA KG X SOL", "prod_ventakgxsol"),
            ]:
                v = dax_productividad_pauno(token, ws_id, ids["productividad_ds"], m_name, label)
                if v is not None:
                    scanned.setdefault("productividad_ds", {})[m_name] = v
                    print(f"    ✓ Productividad [{m_name}] filtrado = {v}")
                else:
                    print(f"    ✗ Productividad [{m_name}] — la consulta filtrada no devolvió valor")

            # Venta Neta (KG): no es medida, es columna SUM con 2 filtros propios
            # (CATEGORIZACION="VENTA BRUTA" + TIPO DE NEGOCIO N2) — ver dax_productividad_venta_neta_kg
            v_vn = dax_productividad_venta_neta_kg(token, ws_id, ids["productividad_ds"])
            if v_vn is not None:
                scanned.setdefault("productividad_ds", {})["__VENTA_NETA_KG__"] = v_vn
                print(f"    ✓ Productividad [Venta Neta KG] filtrado = {v_vn}")
            else:
                print("    ✗ Productividad [Venta Neta KG] — la consulta filtrada no devolvió valor")

        # ── Productividad: schema REST + probe medidas reales (único que no es Live Connection)
        if "productividad_ds" in ids and not scanned.get("productividad_ds"):
            print("  Productividad: leyendo schema REST...")
            schema_prod = get_dataset_schema(token, ws_id, ids["productividad_ds"])
            if schema_prod:
                print(f"    Tablas/medidas Prod: { {t: s['measures'][:5] for t,s in schema_prod.items()} }")
                for tbl_name, tbl_info in schema_prod.items():
                    for m_name in tbl_info["measures"]:
                        v, exists = try_measure(token, ws_id, ids["productividad_ds"], m_name)
                        if exists:
                            scanned.setdefault("productividad_ds", {})[m_name] = v
                            print(f"      ✓ Prod [{m_name}] = {v}")
                for tbl_name, tbl_info in schema_prod.items():
                    if tbl_name.lower() in ["calendario", "calendar", "fecha", "medidas"]:
                        continue
                    for col_name in tbl_info["columns"]:
                        col_lower = col_name.lower()
                        if any(kw in col_lower for kw in ["planilla", "monto", "costo", "kg", "produccion", "total"]):
                            rows_s = dax(token, ws_id, ids["productividad_ds"],
                                         f"EVALUATE ROW(\"v\", SUM('{tbl_name}'[{col_name}]))",
                                         f"prod_sum")
                            if rows_s:
                                sv = rows_s[0].get("[v]") or rows_s[0].get("v")
                                if sv is not None and float(sv or 0) != 0:
                                    lbl = "Planilla Total" if any(k in col_lower for k in ["planilla","costo","monto"]) else "Kg Producidos"
                                    scanned.setdefault("productividad_ds", {})[lbl] = sv
                                    print(f"    Prod SUM '{tbl_name}'[{col_name}] = {sv}")

        # Guardar caché de medidas confirmadas — merge con caché anterior (no borrar)
        existing_cache = load_measure_cache()
        for ds_key, found in scanned.items():
            if found:  # solo actualizar si encontramos algo — preservar entradas de runs previos
                existing_cache[ds_key] = {k: None for k in found.keys()}
        save_measure_cache(existing_cache)

        # Merge consumo→compras: tomar medidas del dataset consumo para calcular ratio
        if scanned.get("consumo"):
            cons = scanned["consumo"]
            # Consumo mensual
            consumo_v = cons.get("Consumo")
            if consumo_v is not None:
                scanned.setdefault("compras", {})["Consumo"] = consumo_v
            # Compras mensual desde dataset consumo (si es razonable < S/20M)
            compras_v = cons.get("Compras") or cons.get("Total Compras")
            if compras_v is not None:
                try:
                    if abs(float(compras_v)) < 20_000_000:
                        scanned.setdefault("compras", {})["Valor Compras"] = compras_v
                        print(f"    Compras desde consumo dataset: {compras_v}")
                except: pass
            # Ratio directo desde dataset consumo
            ratio_v = cons.get("Ratio") or cons.get("% Ratio")
            if ratio_v is not None:
                scanned.setdefault("compras", {})["Ratio"] = ratio_v
                print(f"    Ratio desde consumo dataset: {ratio_v}")

        # Combinamos planificacion con inventario (avance vs ppto)
        if scanned.get("planificacion"):
            scanned.setdefault("inventario", {}).update(scanned["planificacion"])

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
            dias_str = f"{dias:.0f}" if dias is not None else "—"
            print(f"  CxP dias={dias_str} sem={r['estado']}")
            if r["estado"] == "red" and dias:
                summary["agenda_ceo"]["escalar_semana"].append({
                    "empresa": empresa, "texto": f"CxP {dias:.0f}d — riesgo con proveedores.", "responsable": "Finanzas"
                })

        # ── Margen
        if scanned.get("margen"):
            margen_ds_id_total = DATASET_IDS.get(empresa, {}).get("margen")
            if margen_ds_id_total:
                # Sobrescribe Precio x Kilo / Costo x Kilo con el filtro exacto de
                # 12 condiciones confirmado con Copiar consulta (2026-09-04) — el
                # escaneo genérico solo aplicaba filtro de fecha, sin las 11
                # exclusiones adicionales del reporte "Análisis de Ventas".
                p_tot, c_tot = dax_margen_total(token, ws_id, margen_ds_id_total, PREV_YEAR)
                if p_tot is not None and c_tot is not None:
                    scanned["margen"]["Precio x Kilo"] = p_tot
                    scanned["margen"]["Costo x Kilo"] = c_tot
                    print(f"    ✓ Margen total filtrado: precio=S/{p_tot:.2f} costo=S/{c_tot:.2f}")
                else:
                    print("    ✗ Margen total filtrado — la consulta no devolvió valor")
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

            # ── Margen por UEN (B&D, MAQUILA, TIGO) — filtro exacto confirmado
            # con Copiar consulta el 2026-09-04 (ver dax_margen_uen). Reemplaza el
            # bloque anterior que adivinaba nombres de tabla ("Maestro Productos",
            # "Productos", "DimProducto") y nunca encontraba dato — la tabla real
            # es 'Exl Tipo de Negocio'[TIPO DE NEGOCIO N1].
            margen_ds_id = DATASET_IDS.get(empresa, {}).get("margen")
            if margen_ds_id:
                uen_data = []
                for uen in ["B&D", "MAQUILA", "TIGO"]:
                    precio, costo = dax_margen_uen(token, ws_id, margen_ds_id, uen, PREV_YEAR, label=f"margen-uen-{uen}")
                    if precio and costo and precio > 0:
                        pct = (precio - costo) / precio * 100
                        uen_data.append({
                            "uen": uen, "margen": f"{pct:.1f}%",
                            "precio_kg": f"S/{precio:.2f}", "costo_kg": f"S/{costo:.2f}",
                            "estado": "green" if pct >= 52 else ("yellow" if pct >= 46 else "red"),
                        })
                        print(f"    ✓ Margen UEN [{uen}]: precio=S/{precio:.2f} costo=S/{costo:.2f} margen={pct:.1f}%")
                    else:
                        print(f"    ✗ Margen UEN [{uen}] — la consulta filtrada no devolvió valor")
                if uen_data:
                    empresa_data["reportes"]["margen_variable"]["por_uen"] = uen_data
                    print(f"  Margen UEN: {uen_data}")

        # ── Consumo de Materiales (PRODUCCIÓN: CECO Ajustado + MIP/TN)
        if scanned.get("consumo"):
            rc = build_consumo_materiales(scanned["consumo"])
            if rc:
                empresa_data["reportes"]["consumo_materiales"] = rc
                print(f"  Consumo Materiales {len(rc['kpis'])} KPIs sem={rc['estado']}")
            else:
                # Live Connection: guardar marcador para que el frontend muestre mensaje correcto
                empresa_data["reportes"]["consumo_materiales"] = {
                    "estado": "grey", "alerta": None, "kpis": [],
                    "live_connection": True
                }
                print("  Consumo Materiales: sin medidas DAX — Live Connection")

        # ── Mermas
        if scanned.get("mermas"):
            r = build_mermas(scanned["mermas"])
            empresa_data["reportes"]["mermas"] = r
            print(f"  Mermas {len(r['kpis'])} KPIs sem={r['estado']}")

            # ── Mermas por UEN — filtro exacto confirmado con Copiar consulta (2026-09-04)
            # Reemplaza el bloque anterior que adivinaba tabla/columna ("Maestro Productos",
            # "UEN", "Unidad Negocio"...) y nunca encontraba dato real.
            mermas_ds_id = DATASET_IDS.get(empresa, {}).get("mermas")
            if mermas_ds_id:
                # (medida en 'Tabla Mermas', filtro extra propio de la tarjeta o None)
                UEN_MERMA_CONFIG = [
                    ("B&D",     "% Merma total B&D",     None),
                    ("TIGO",    "% Merma total TIGO",
                     "FILTER(\n      KEEPFILTERS(VALUES('Tabla Mermas'[TIPO DE BASE])),\n      NOT('Tabla Mermas'[TIPO DE BASE] IN {BLANK()})\n    )"),
                    # Nombre real confirmado con doble M: "MMAQUILA", no "MAQUILA"
                    ("MAQUILA", "% Merma total MMAQUILA",
                     "FILTER(\n      KEEPFILTERS(VALUES('Tabla Mermas'[TIPO DE BASE])),\n      NOT('Tabla Mermas'[TIPO DE BASE] IN {BLANK()})\n    )"),
                ]
                uen_merma = []
                for uen, medida, extra in UEN_MERMA_CONFIG:
                    v = dax_mermas_uen(token, ws_id, mermas_ds_id, medida, PREV_YEAR,
                                        extra_filtro=extra, label=f"merma-uen-{uen}")
                    if v is not None:
                        if abs(v) < 1: v *= 100
                        s = "red" if v > 3 else ("yellow" if v > 2 else "green")
                        uen_merma.append({"uen": uen, "merma": f"{v:.2f}%", "estado": s})
                        print(f"    ✓ Merma UEN [{uen}] filtrado = {v:.2f}%")
                    else:
                        print(f"    ✗ Merma UEN [{uen}] — la consulta filtrada no devolvió valor")
                if uen_merma:
                    empresa_data["reportes"]["mermas"]["por_uen"] = uen_merma
                    print(f"  Mermas UEN: {uen_merma}")

                # Por Planta (ATE / PACHACAMAC / TERCEROS) — filtro exacto confirmado con
                # Copiar consulta el 2026-09-04, capturado desde el gráfico "Merma Mensual
                # por Planta" (pestaña RESUMEN). En el modelo, cada planta corresponde a un
                # almacén con nombre propio (no "ATE"/"PACHACAMAC"/"TERCEROS" literal):
                #   Ate         -> almacén "Lácteos Producción"     · medida genérica
                #   Pachacámac  -> almacén "Salsas Producción"       · medida "...SALSAS"
                #   Terceros    -> almacenes de maquiladores externos (Abuela Maquila,
                #                  Piamonte Maquila, Lácteos Dosimetría) · medida "...MAQUILA"
                #                  (una sola M — distinta de la medida UEN "MMAQUILA")
                # (planta, medida, almacenes, con_tipo_base) — Terceros NO trae el
                # filtro TIPO DE BASE, confirmado ausente en su Copiar consulta.
                PLANTA_MERMA_CONFIG = [
                    ("ATE",        "% Merma total",         ["Lacteos Producción"], True),
                    ("PACHACAMAC", "% Merma total SALSAS",  ["Salsas Producción"], True),
                    ("TERCEROS",   "% Merma total MAQUILA", ["Abuela Maquila", "Piamonte Maquila", "Lacteos Dosimetria"], False),
                ]
                planta_merma = []
                for planta, medida, almacenes, con_tb in PLANTA_MERMA_CONFIG:
                    v = dax_mermas_planta(token, ws_id, mermas_ds_id, medida, almacenes, PREV_YEAR,
                                           con_tipo_base=con_tb, label=f"merma-planta-{planta}")
                    if v is not None:
                        if abs(v) < 1: v *= 100
                        s = "red" if v > 3 else ("yellow" if v > 2 else "green")
                        planta_merma.append({"planta": planta, "merma": f"{v:.2f}%", "estado": s})
                        print(f"    ✓ Merma Planta [{planta}] filtrado = {v:.2f}%")
                    else:
                        print(f"    ✗ Merma Planta [{planta}] — la consulta filtrada no devolvió valor")
                if planta_merma:
                    empresa_data["reportes"]["mermas"]["por_planta"] = planta_merma
                    print(f"  Mermas Planta: {planta_merma}")

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

        # ── Inventario — filtro exacto confirmado con Copiar consulta (2026-09-04)
        # sobre el reporte '6. Rotación de inventario' (página RI Clasificación).
        inv_ds_id = DATASET_IDS.get(empresa, {}).get("inventario")
        if inv_ds_id:
            cob_total = dax_inventario_kardex(token, ws_id, inv_ds_id, "Días Rotación", PREV_YEAR, PREV_MONTH)
            cob_mp    = dax_inventario_kardex(token, ws_id, inv_ds_id, "Días Rotación MP", PREV_YEAR, PREV_MONTH)
            if cob_total is not None:
                scanned.setdefault("inventario", {})["Cobertura Total (días)"] = cob_total
                print(f"    ✓ Inventario [Cobertura Total] filtrado = {cob_total:.1f} días")
            if cob_mp is not None:
                scanned.setdefault("inventario", {})["Cobertura MP (días)"] = cob_mp
                print(f"    ✓ Inventario [Cobertura MP] filtrado = {cob_mp:.1f} días")

            composicion = dax_inventario_composicion(token, ws_id, inv_ds_id)
            if composicion:
                scanned.setdefault("inventario", {})["_composicion"] = composicion
                print(f"    ✓ Inventario [Composición]: {composicion}")

        # ── Inventario
        if scanned.get("inventario"):
            r = build_inventario(scanned["inventario"])
            empresa_data["reportes"]["sop_inventario"] = r
            print(f"  S&OP {len(r['kpis'])} KPIs sem={r['estado']}")

        # ── Control Interno — filtro exacto confirmado con Copiar consulta (2026-09-04)
        # sobre '8. Reporte de auditoria' (Dashboard de Control Interno): solo
        # Empresa="Pauno", sin filtro de fecha (acumulado histórico total).
        control_ds_id = DATASET_IDS.get(empresa, {}).get("control_ds")
        if control_ds_id:
            satisf_v = dax_control_interno(token, ws_id, control_ds_id, "Satisfactorio")
            obs_v    = dax_control_interno(token, ws_id, control_ds_id, "Con Observaciones")
            crit_v   = dax_control_interno(token, ws_id, control_ds_id, "Critico")
            if satisf_v is not None:
                scanned.setdefault("control_ds", {})["Satisfactorio"] = satisf_v
            if obs_v is not None:
                scanned.setdefault("control_ds", {})["Con Observaciones"] = obs_v
            if crit_v is not None:
                scanned.setdefault("control_ds", {})["Critico"] = crit_v
            if satisf_v is not None or obs_v is not None or crit_v is not None:
                print(f"    ✓ Control Interno: Satisfactorio={satisf_v} Con Observaciones={obs_v} Critico={crit_v}")

            planes_abiertos_v = dax_planes_accion(token, ws_id, control_ds_id, "Planes Abiertos")
            if planes_abiertos_v is not None:
                scanned.setdefault("control_ds", {})["Planes Abiertos"] = planes_abiertos_v
                print(f"    ✓ Control Interno: Planes Abiertos={planes_abiertos_v}")

            planes_cerrados_v = dax_planes_accion(token, ws_id, control_ds_id, "Planes Cerrados")
            if planes_cerrados_v is not None:
                scanned.setdefault("control_ds", {})["Planes Cerrados"] = planes_cerrados_v
                print(f"    ✓ Control Interno: Planes Cerrados={planes_cerrados_v}")

            total_planes_v = dax_planes_accion(token, ws_id, control_ds_id, "Total Planes de Acción")
            if total_planes_v is not None:
                scanned.setdefault("control_ds", {})["Total Planes de Acción"] = total_planes_v
                print(f"    ✓ Control Interno: Total Planes de Acción={total_planes_v}")

            planes_atrasados_v = dax_planes_accion(token, ws_id, control_ds_id, "Planes Atrasados")
            if planes_atrasados_v is not None:
                scanned.setdefault("control_ds", {})["Planes Atrasados"] = planes_atrasados_v
                print(f"    ✓ Control Interno: Planes Atrasados={planes_atrasados_v}")

            ejecutado_v = dax_control_interno_sum(token, ws_id, control_ds_id, "¿Ejecutado?")
            if ejecutado_v is not None:
                scanned.setdefault("control_ds", {})["Ejecutado"] = ejecutado_v
                print(f"    ✓ Control Interno: Ejecutado (suma)={ejecutado_v}")

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

        # ── Fill Rate (dataset dedicado 12. Calculo de Provisiones)
        if scanned.get("fill_rate"):
            fr = build_fill_rate(scanned["fill_rate"])
            if fr["kpis"]:
                empresa_data["reportes"]["fill_rate"] = fr
                print(f"  Fill Rate {fr['kpis'][0]['valor'] if fr['kpis'] else '—'}")

        # ── Avance vs Presupuesto (planificacion mergeado en inventario)
        if scanned.get("inventario"):
            av, avance_pct = build_avance(scanned["inventario"])
            if av["kpis"]:
                empresa_data["reportes"]["margen_variable_pag2"] = av
                print(f"  Avance PPTO {fmt_pct(avance_pct)} sem={av['estado']}")

        summary["empresas"][empresa] = empresa_data

    if holding_ventas: summary["holding_ventas"] = fmt_soles(holding_ventas)
    if holding_mora_vals:
        summary["holding_mora"] = f"{sum(holding_mora_vals)/len(holding_mora_vals):.1f}%"

    # Inyectar datos YoY del discover si existen
    discovered_path = OUTPUT_DIR / "discovered_measures.json"
    if discovered_path.exists():
        try:
            disc = json.loads(discovered_path.read_text())
            yoy  = disc.get("yoy", {})
            comp = disc.get("comp_filtro", "")
            if yoy:
                summary["yoy_periodo"] = comp
                summary["yoy"] = {}
                # Ventas año anterior
                ventas_yoy = (yoy.get("margen", {}) or {}).get("Venta Total")
                if ventas_yoy is not None:
                    summary["yoy"]["ventas"] = ventas_yoy
                # Fill Rate año anterior
                fr_yoy = (yoy.get("fill_rate", {}) or {}).get("% Fill Rate")
                if fr_yoy is not None:
                    summary["yoy"]["fill_rate"] = fr_yoy
                # Mermas año anterior
                mermas_yoy = (yoy.get("mermas", {}) or {}).get("% Merma Total")
                if mermas_yoy is not None:
                    summary["yoy"]["mermas"] = mermas_yoy
                # Avance año anterior
                avance_yoy = (yoy.get("planificacion", {}) or {}).get("% Avance")
                if avance_yoy is not None:
                    summary["yoy"]["avance"] = avance_yoy
                # Morosidad año anterior
                mora_yoy = (yoy.get("cxc", {}) or {}).get("% Morosidad") or \
                           (yoy.get("cxc", {}) or {}).get("Morosidad")
                if mora_yoy is not None:
                    summary["yoy"]["morosidad"] = mora_yoy
                print(f"  YoY {comp}: ventas={ventas_yoy} fr={fr_yoy} mermas={mermas_yoy}")
        except Exception as e:
            print(f"  YoY no disponible: {e}")

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
