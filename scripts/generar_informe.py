"""Genera el informe del dashboard como PDF mediante un navegador sin interfaz."""

import os
from datetime import datetime
from pathlib import Path

from playwright.sync_api import sync_playwright


output_dir = Path("informes")
output_dir.mkdir(exist_ok=True)
commit = os.environ.get("GITHUB_SHA", "local")[:7]
timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
output_file = output_dir / f"informe_mantencion_{timestamp}_{commit}.pdf"

with sync_playwright() as playwright:
    browser = playwright.chromium.launch(headless=True)
    context = browser.new_context(
        locale="es-CL",
        timezone_id="America/Santiago",
        viewport={"width": 1440, "height": 1000},
    )
    page = context.new_page()
    page.goto("http://127.0.0.1:8000/index.html", wait_until="domcontentloaded", timeout=120_000)
    page.wait_for_function(
        """() => {
          const estado = document.getElementById('estadoCarga')?.innerText || '';
          return estado.includes('registros leídos') && !estado.includes('No se pudieron cargar');
        }""",
        timeout=180_000,
    )
    page.evaluate("() => generarInformePDF(false, ['resumen', 'diario', 'plan'])")
    page.wait_for_selector("#printReport .pdfPage", state="attached", timeout=60_000)
    page.emulate_media(media="print")
    page.pdf(
        path=str(output_file),
        format="Letter",
        landscape=False,
        print_background=True,
        prefer_css_page_size=True,
        margin={"top": "0", "right": "0", "bottom": "0", "left": "0"},
    )
    browser.close()

if output_file.stat().st_size < 10_000:
    raise RuntimeError(f"El PDF generado parece incompleto: {output_file}")

print(output_file.as_posix())
