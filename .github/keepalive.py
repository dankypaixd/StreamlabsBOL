"""Abre la app como lo haria una persona para que Streamlit Cloud no la duerma.

Streamlit Community Cloud duerme las apps que pasan unas 12 horas sin visitas. Un simple ping HTTP no
cuenta como visita (la app se dibuja por websocket), por eso aqui se abre un navegador de verdad.
Si la app ya estaba dormida, pulsa el boton para despertarla.

La direccion sale de la variable APP_URL del repositorio (Settings > Secrets and variables > Actions >
Variables) o, si no existe, del archivo .github/app_url.txt.
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

from playwright.sync_api import TimeoutError as PlaywrightTimeout
from playwright.sync_api import sync_playwright

WAKE_BUTTON = "Yes, get this app back up!"


def app_url() -> str:
    url = os.environ.get("APP_URL", "").strip()
    if not url:
        path = Path(__file__).with_name("app_url.txt")
        url = path.read_text(encoding="utf-8").strip() if path.exists() else ""
    return url


def main() -> int:
    url = app_url()
    if not url.startswith("http"):
        print("Falta la direccion de la app: crea la variable APP_URL o el archivo .github/app_url.txt")
        return 1
    with sync_playwright() as p:
        options = {"executable_path": os.environ["CHROME_PATH"]} if os.environ.get("CHROME_PATH") else {}
        browser = p.chromium.launch(args=["--no-sandbox"], **options)
        page = browser.new_page()
        page.goto(url, wait_until="domcontentloaded", timeout=60_000)
        page.wait_for_timeout(8_000)
        wake = page.get_by_role("button", name=WAKE_BUTTON)
        if wake.count():
            print("La app estaba dormida: la despierto.")
            wake.first.click()
        try:
            page.get_by_role("tab").first.wait_for(timeout=240_000)  # la app dibujo sus pestanas
        except PlaywrightTimeout:
            print("La app no cargo a tiempo.")
            browser.close()
            return 1
        page.wait_for_timeout(15_000)  # unos segundos conectada para que cuente como visita
        print("La app esta despierta.")
        browser.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
