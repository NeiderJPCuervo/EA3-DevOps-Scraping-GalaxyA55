"""Toma las capturas de pantalla del informe de la EA3 con Selenium (Firefox headless).

Solo visita páginas públicas: el repositorio y sus ejecuciones en GitHub y la app
desplegada en Render. Espera a que cada página termine de dibujarse (diagramas de
los workflows, imágenes diferidas) antes de capturarla.

Uso (desde la raíz del repositorio):
    python informe/capturar_pantallas.py
"""
import json
import subprocess
import time
from pathlib import Path

from selenium import webdriver
from selenium.webdriver.firefox.options import Options

RAIZ = Path(__file__).resolve().parent.parent
CAPTURAS = RAIZ / "informe" / "capturas_ea3"
REPO = "NeiderJPCuervo/EA3-DevOps-Scraping-GalaxyA55"
GH = f"https://github.com/{REPO}"
APP = "https://ea3-galaxy-a55.onrender.com"


def ultima_ejecucion(workflow=None, rama=None):
    cmd = ["gh", "run", "list", "-R", REPO, "-L", "1", "--json", "databaseId"]
    if workflow:
        cmd += ["-w", workflow]
    if rama:
        cmd += ["-b", rama]
    return json.loads(subprocess.run(cmd, check=True, capture_output=True, text=True).stdout)[0]["databaseId"]


def capturar(driver, nombre, url, alto=1000, espera=4, desplazar=False):
    driver.set_window_size(1366, alto)
    driver.get(url)
    time.sleep(espera)
    if desplazar:   # fuerza la carga de las imágenes con loading="lazy"
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight)")
        time.sleep(2)
        driver.execute_script("window.scrollTo(0, 0)")
        time.sleep(1)
    driver.save_screenshot(str(CAPTURAS / nombre))
    print(f"{nombre}: {url}")


def main():
    CAPTURAS.mkdir(parents=True, exist_ok=True)
    opciones = Options()
    opciones.add_argument("--headless")
    opciones.set_preference("intl.accept_languages", "es-CO,es")
    driver = webdriver.Firefox(options=opciones)
    try:
        capturar(driver, "01_repo.png", GH, 1100)
        capturar(driver, "02_actions.png", f"{GH}/actions", 1000)
        capturar(driver, "03_pr6_checks.png", f"{GH}/pull/6/checks", 900)
        capturar(driver, "04_ci_main.png", f"{GH}/actions/runs/{ultima_ejecucion('CI', 'main')}", 1100, espera=6)
        capturar(driver, "05_deploy_render.png", f"{GH}/actions/runs/{ultima_ejecucion('Deploy Render', 'main')}",
                 1250, espera=6)
        capturar(driver, "06_pr11_checks.png", f"{GH}/pull/11/checks", 900)
        capturar(driver, "07_ci_fallido.png", f"{GH}/actions/runs/{ultima_ejecucion(rama='demo/imagen-sin-datos')}",
                 1150, espera=6)
        capturar(driver, "08_issues.png", f"{GH}/issues?q=is%3Aissue", 900)
        capturar(driver, "09_pulls.png", f"{GH}/pulls?q=is%3Apr", 1000)
        capturar(driver, "10_app_tablero.png", APP, 1500, espera=8, desplazar=True)
        capturar(driver, "11_app_docs.png", f"{APP}/docs", 1050, espera=6)
        capturar(driver, "12_app_salud.png", f"{APP}/salud", 500, espera=3)
    finally:
        driver.quit()


if __name__ == "__main__":
    main()
