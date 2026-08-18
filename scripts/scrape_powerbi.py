"""
scrape_powerbi.py — Juanito Garcia
Captura dashboards Power BI en paralelo (4 páginas simultáneas).
Compatible con GitHub Actions (headless) y Mac local.
"""

import asyncio
import sys
import os
import argparse
import shutil
from datetime import datetime
from pathlib import Path

try:
    from playwright.async_api import async_playwright
except ImportError:
    print("ERROR: pip install playwright && python -m playwright install chromium")
    sys.exit(1)

EMPRESAS = {
    "PAUNO": [
        {"id":  1, "nombre": "cuentas_por_cobrar",   "url": "https://app.powerbi.com/view?r=eyJrIjoiNmMxNjNmZGQtMjUxYS00MzhkLTkyYmQtNmVlMWU2MmU2MGY0IiwidCI6Ijc3ZTAwODA5LWEzNmUtNDkzNS04ZWRjLTU1YWI3Y2YxODYxZiIsImMiOjR9"},
        {"id":  2, "nombre": "cuentas_por_pagar",    "url": "https://app.powerbi.com/view?r=eyJrIjoiZDQzM2QyNzQtODIxYS00ZTk0LThmZDQtZWQyMjE0Mzg5MGNmIiwidCI6Ijc3ZTAwODA5LWEzNmUtNDkzNS04ZWRjLTU1YWI3Y2YxODYxZiIsImMiOjR9"},
        {"id":  3, "nombre": "margen_variable",      "url": "https://app.powerbi.com/view?r=eyJrIjoiZWY4YWViZWItMGI4OS00ZDc5LWE5YmItMjI5ODExNmE4MTVlIiwidCI6Ijc3ZTAwODA5LWEzNmUtNDkzNS04ZWRjLTU1YWI3Y2YxODYxZiIsImMiOjR9"},
        {"id":  4, "nombre": "mermas",               "url": "https://app.powerbi.com/view?r=eyJrIjoiMTFjMTZmYWUtNDEyOS00ZDkyLWJkYTYtZDhiYzY0ZjI3MzE2IiwidCI6Ijc3ZTAwODA5LWEzNmUtNDkzNS04ZWRjLTU1YWI3Y2YxODYxZiIsImMiOjR9"},
        {"id":  5, "nombre": "compras",              "url": "https://app.powerbi.com/view?r=eyJrIjoiYTFiZWRkZGUtNGY0NC00ODZmLWEwNzktYTE1ZDBjNTFjODQ5IiwidCI6Ijc3ZTAwODA5LWEzNmUtNDkzNS04ZWRjLTU1YWI3Y2YxODYxZiIsImMiOjR9"},
        {"id":  6, "nombre": "dias_inventario",      "url": "https://app.powerbi.com/view?r=eyJrIjoiOGMxMmJkODItYWFjNi00NjdiLWFkYWItMjNhYTU3NTkwZTNhIiwidCI6Ijc3ZTAwODA5LWEzNmUtNDkzNS04ZWRjLTU1YWI3Y2YxODYxZiIsImMiOjR9"},
        {"id":  7, "nombre": "control_interno",      "url": "https://app.powerbi.com/view?r=eyJrIjoiZDU3MzhiYjgtOTgyYi00YjllLTkwOGQtZjg5MTM4ZDFlNWMwIiwidCI6Ijc3ZTAwODA5LWEzNmUtNDkzNS04ZWRjLTU1YWI3Y2YxODYxZiIsImMiOjR9"},
        {"id":  8, "nombre": "consumo_materiales",   "url": "https://app.powerbi.com/view?r=eyJrIjoiZDhhY2M1NjEtMTU1Mi00M2M4LWEzNmQtYzk4ZjhlYjZlZTg5IiwidCI6Ijc3ZTAwODA5LWEzNmUtNDkzNS04ZWRjLTU1YWI3Y2YxODYxZiIsImMiOjR9"},
        {"id":  9, "nombre": "cuentas_contables",    "url": "https://app.powerbi.com/view?r=eyJrIjoiMTg2MDZhMDYtODZmYi00MjNhLWEyNzYtNDA1MTQyYjQ4OTkwIiwidCI6Ijc3ZTAwODA5LWEzNmUtNDkzNS04ZWRjLTU1YWI3Y2YxODYxZiIsImMiOjR9"},
        {"id": 10, "nombre": "sop_inventario",       "url": "https://app.powerbi.com/view?r=eyJrIjoiMmVhNjliNmQtNTU2NC00MzQ0LTljZTAtYzgyMGMwM2Q4Yjg5IiwidCI6Ijc3ZTAwODA5LWEzNmUtNDkzNS04ZWRjLTU1YWI3Y2YxODYxZiIsImMiOjR9"},
        {"id": 11, "nombre": "fill_rate",            "url": "https://app.powerbi.com/view?r=eyJrIjoiMjg4NzRiOTctZGQxOC00NjlmLWIxMTQtMjdkMGE1MTZlYmFiIiwidCI6Ijc3ZTAwODA5LWEzNmUtNDkzNS04ZWRjLTU1YWI3Y2YxODYxZiIsImMiOjR9"},
        {"id": 12, "nombre": "productividad",        "url": "https://app.powerbi.com/view?r=eyJrIjoiNjEzNzRmMTMtOWVmNC00YTk1LTkwMzctZDFhNDRjZWE5NTRkIiwidCI6Ijc3ZTAwODA5LWEzNmUtNDkzNS04ZWRjLTU1YWI3Y2YxODYxZiIsImMiOjR9"},
        {"id": 13, "nombre": "margen_variable_pag2", "url": "https://app.powerbi.com/view?r=eyJrIjoiZWY4YWViZWItMGI4OS00ZDc5LWE5YmItMjI5ODExNmE4MTVlIiwidCI6Ijc3ZTAwODA5LWEzNmUtNDkzNS04ZWRjLTU1YWI3Y2YxODYxZiIsImMiOjR9&pageName=0fa2166e084a7d814807"},
    ],
    "AMAUTA": [
        {"id": 1, "nombre": "reporte_01", "url": "https://app.powerbi.com/view?r=eyJrIjoiMDk2MGY0ZGYtMzJkMC00OTAxLWE2NDAtYmU5OWZhNDU2YjNhIiwidCI6Ijc3ZTAwODA5LWEzNmUtNDkzNS04ZWRjLTU1YWI3Y2YxODYxZiIsImMiOjR9"},
        {"id": 2, "nombre": "reporte_02", "url": "https://app.powerbi.com/view?r=eyJrIjoiYmU1MGEzZDItZWJmNS00YTVjLTk4NDYtZGQyYTE5YmE2NTlhIiwidCI6Ijc3ZTAwODA5LWEzNmUtNDkzNS04ZWRjLTU1YWI3Y2YxODYxZiIsImMiOjR9"},
        {"id": 3, "nombre": "reporte_03", "url": "https://app.powerbi.com/view?r=eyJrIjoiODU0YmJhMjUtMTg5Ni00ODcwLTg3MzEtYTQ3NWVkM2Q1OTFhIiwidCI6Ijc3ZTAwODA5LWEzNmUtNDkzNS04ZWRjLTU1YWI3Y2YxODYxZiIsImMiOjR9"},
    ],
    "NEOPACK": [
        {"id": 1, "nombre": "reporte_01", "url": "https://app.powerbi.com/view?r=eyJrIjoiZDZiMTRiM2EtNzZhZi00NmMzLTllNmEtMGNiMjYxY2UzYTg4IiwidCI6Ijc3ZTAwODA5LWEzNmUtNDkzNS04ZWRjLTU1YWI3Y2YxODYxZiIsImMiOjR9"},
        {"id": 2, "nombre": "reporte_02", "url": "https://app.powerbi.com/view?r=eyJrIjoiNzRiZGE4ZGItZWE1Mi00YjllLWEzMGMtOWU2MTgyMDUwNDI2IiwidCI6Ijc3ZTAwODA5LWEzNmUtNDkzNS04ZWRjLTU1YWI3Y2YxODYxZiIsImMiOjR9"},
        {"id": 3, "nombre": "reporte_03", "url": "https://app.powerbi.com/view?r=eyJrIjoiNjE2YWEzMzYtZTQyMy00NDM5LTk0NTAtODUxMDU2Njk2MTY5IiwidCI6Ijc3ZTAwODA5LWEzNmUtNDkzNS04ZWRjLTU1YWI3Y2YxODYxZiIsImMiOjR9"},
    ],
}

WAIT_MS = 15000
WORKERS = 4
MAX_RETRIES = 2


async def capturar_reporte(page, empresa, reporte, carpeta_empresa, intento=1):
    nombre_archivo = f"{reporte['id']:02d}_{reporte['nombre']}.png"
    ruta_salida = carpeta_empresa / nombre_archivo
    tag = f"[{empresa}/{reporte['nombre']}]"
    try:
        await page.goto(reporte["url"], wait_until="load", timeout=45000)
        await page.wait_for_timeout(WAIT_MS)
        await page.screenshot(path=str(ruta_salida), full_page=False)
        print(f"  OK  {tag}")
        return True
    except Exception as e:
        if intento <= MAX_RETRIES:
            print(f"  RETRY {intento} {tag} ({str(e)[:50]})")
            await page.wait_for_timeout(5000)
            return await capturar_reporte(page, empresa, reporte, carpeta_empresa, intento + 1)
        print(f"  FALLO {tag} ({str(e)[:50]})")
        return False


async def worker(worker_id, tareas, carpeta_salida, context, resultados):
    page = await context.new_page()
    for empresa, reporte in tareas:
        carpeta_empresa = carpeta_salida / empresa
        ok = await capturar_reporte(page, empresa, reporte, carpeta_empresa)
        resultados.append(ok)
    await page.close()


async def main(output_dir: str):
    carpeta_salida = Path(output_dir)

    # Limpiar carpeta anterior y recrear
    if carpeta_salida.exists():
        shutil.rmtree(carpeta_salida)
    carpeta_salida.mkdir(parents=True)

    for empresa in EMPRESAS:
        (carpeta_salida / empresa).mkdir(exist_ok=True)

    tareas = [(empresa, r) for empresa, reportes in EMPRESAS.items() for r in reportes]
    total = len(tareas)

    print(f"Juanito Garcia — scraping paralelo ({WORKERS} workers)")
    print(f"Reportes: {total} | Espera por reporte: {WAIT_MS//1000}s")
    t0 = datetime.now()

    async with async_playwright() as pw:
        browser = await pw.chromium.launch(
            headless=True,
            args=["--no-sandbox", "--disable-dev-shm-usage", "--disable-gpu"]
        )
        context = await browser.new_context(
            viewport={"width": 1920, "height": 1080},
            locale="es-PE"
        )

        grupos = [tareas[i::WORKERS] for i in range(WORKERS)]
        resultados = []
        await asyncio.gather(*[
            worker(i, grupos[i], carpeta_salida, context, resultados)
            for i in range(WORKERS)
        ])

        await browser.close()

    total_ok = sum(resultados)
    elapsed = (datetime.now() - t0).seconds
    print(f"\nTotal: {total_ok}/{total} OK en {elapsed}s ({elapsed//60}m {elapsed%60}s)")
    return str(carpeta_salida)


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--output", default="data/latest")
    args = parser.parse_args()
    asyncio.run(main(args.output))
