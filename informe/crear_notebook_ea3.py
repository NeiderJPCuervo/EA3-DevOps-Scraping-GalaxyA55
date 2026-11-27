"""Construye el notebook de la EA3 (Cuervo_Quintero_John_Actividad_EA3.ipynb).

El notebook documenta y ejecuta paso a paso la app web, las pruebas, la imagen
Docker, el pipeline de GitHub Actions y la verificación de producción en Render.
Lee informe/datos_ea3.json (recolectar_datos_ea3.py).

Uso (desde la raíz del repositorio):
    python informe/crear_notebook_ea3.py
    jupyter nbconvert --to notebook --execute --inplace informe/Cuervo_Quintero_John_Actividad_EA3.ipynb
"""
from pathlib import Path

import nbformat as nbf

DESTINO = Path(__file__).resolve().parent / "Cuervo_Quintero_John_Actividad_EA3.ipynb"

celdas = []


def md(texto):
    celdas.append(nbf.v4.new_markdown_cell(texto.strip()))


def code(texto):
    celdas.append(nbf.v4.new_code_cell(texto.strip()))


md("""
# App web con CI/CD en GitHub Actions y despliegue en Render
## Evidencia de aprendizaje 3 · Samsung Galaxy A55 en Mercado Libre Colombia

**Autor:** John Neider Cuervo Quintero
**Asignatura:** Programación para Análisis de Datos (PREICA2602B020048) · **Docente:** Andrés Felipe Palacio
**Institución Universitaria Digital de Antioquia** · Ingeniería de Software y Datos

| | Enlace |
|---|---|
| Repositorio | https://github.com/NeiderJPCuervo/EA3-DevOps-Scraping-GalaxyA55 |
| App en producción (Render) | https://ea3-galaxy-a55.onrender.com |
| Documentación de la API | https://ea3-galaxy-a55.onrender.com/docs |
| Ejecuciones de GitHub Actions | https://github.com/NeiderJPCuervo/EA3-DevOps-Scraping-GalaxyA55/actions |

Este notebook recorre la evidencia en el mismo orden en que la construí: los datos que heredé de la EA1, la app
web, sus pruebas, la imagen Docker, el pipeline de CI/CD y, al final, la verificación de la app en producción.
Todas las celdas se ejecutaron; las que consultan GitHub y Render usan datos reales del día de la entrega.
""")

md("""
## 1. De dónde viene este proyecto

| Evidencia | Qué construí | Resultado |
|---|---|---|
| EA1 | Scraper con Selenium, BeautifulSoup y Scrapy sobre Mercado Libre | 119 publicaciones, 30 del Galaxy A55 |
| EA2 | Flujo DevOps: Git, GitHub Flow, CI con pytest y CD a GitHub Pages | Tablero estático publicado en cada merge |
| **EA3** | **App FastAPI + Docker + CI/CD con despliegue en Render** | **API y tablero en la nube** |
""")

code("""
import json
import os
import subprocess
import sys
import tempfile
import urllib.request
from pathlib import Path

import pandas as pd

# El notebook funciona desde la raíz del repositorio o desde informe/ / entrega/
RAIZ = Path.cwd()
while not (RAIZ / "app").exists() and RAIZ != RAIZ.parent:
    RAIZ = RAIZ.parent
sys.path[:0] = [str(RAIZ), str(RAIZ / "scraping")]

D = json.loads((RAIZ / "informe" / "datos_ea3.json").read_text(encoding="utf-8"))
resumen = json.loads((RAIZ / "data" / "resumen.json").read_text(encoding="utf-8"))
pd.set_option("display.max_colwidth", 70)
print("Repositorio:", D["url_repo"], "| main en", D["main"])
""")

md("## 2. Los datos que sirve la app")

code("""
def cop(x):
    return f"${x:,.0f}".replace(",", ".")


condicion = {c["condicion"]: c for c in resumen["por_condicion"]}
pd.Series({
    "Publicaciones extraídas": resumen["publicaciones_totales"],
    "Publicaciones del Galaxy A55": resumen["publicaciones_a55"],
    "Precio mediano": cop(resumen["precio"]["50%"]),
    "Mediana nuevo": f"{cop(condicion['Nuevo']['median'])} (n={condicion['Nuevo']['count']})",
    "Mediana usado": f"{cop(condicion['Usado']['median'])} (n={condicion['Usado']['count']})",
    "Con envío gratis": f"{resumen['pct_envio_gratis']:.0f} %",
}, name="Valor").to_frame()
""")

md("""
## 3. La app web (FastAPI)

`app/main.py` reutiliza el código de la EA2: `config.py` para las rutas de los datos y `06_sitio.py` para el
tablero HTML. Aquí la pruebo con el `TestClient` de FastAPI, que ejecuta la app en memoria sin levantar un
servidor.
""")

code("""
import warnings

warnings.filterwarnings("ignore", message=".*httpx.*")   # aviso interno de Starlette
from fastapi.testclient import TestClient

from app.main import app

cliente = TestClient(app)
for ruta in ["/salud", "/api/resumen", "/api/publicaciones?limite=1", "/api/publicaciones/MCO000",
             "/api/publicaciones?precio_max=-1", "/"]:
    r = cliente.get(ruta)
    print(f"{r.status_code}  {r.headers['content-type'].split(';')[0]:<17} {ruta}")
""")

code("""
nuevos = cliente.get("/api/publicaciones", params={"condicion": "Nuevo"}).json()
print("Publicaciones nuevas:", nuevos["total"])
pd.DataFrame(nuevos["publicaciones"])[["id", "titulo", "capacidad_gb", "precio_actual", "envio_gratis"]]
""")

code("""
esquema = cliente.get("/openapi.json").json()
pd.DataFrame([{"Ruta": ruta, "Método": metodo.upper(), "Descripción": op["summary"]}
              for ruta, ops in esquema["paths"].items() for metodo, op in ops.items()])
""")

md("""
## 4. Pruebas automatizadas

La suite de la EA2 (parsers, BeautifulSoup, Scrapy, análisis y tablero) se amplió con `tests/test_app.py`. Se
ejecuta sin Internet sobre el HTML que capturó Selenium y escribe en una carpeta temporal (`EA2_SALIDA`).
""")

code("""
with tempfile.TemporaryDirectory() as tmp:
    r = subprocess.run([sys.executable, "-m", "pytest", "-q", "--cov", "--cov-report=term"], cwd=RAIZ,
                       capture_output=True, text=True, env=os.environ | {"EA2_SALIDA": tmp})
lineas = r.stdout.strip().splitlines()
print("\\n".join(x for x in lineas if x.startswith(("app/", "scraping/", "TOTAL")) or "passed" in x))
""")

code("""
pd.Series(D["pruebas_por_archivo"], name="Pruebas").rename_axis("Archivo").to_frame()
""")

md("""
## 5. La imagen Docker

La misma imagen que corre en Render se construye y se prueba en cada pull request (job `docker` del CI).
""")

code("""
print((RAIZ / "Dockerfile").read_text(encoding="utf-8"))
print(f"Tamaño de la imagen: {D['imagen']['tamano_mb']} MB (base {D['imagen']['base']})")
""")

md("""
## 6. El pipeline de CI/CD

| Workflow | Disparador | Qué hace |
|---|---|---|
| `ci.yml` | PR y push a `main` | ruff, pytest 3.12/3.13, pipeline de datos, build y smoke test de la imagen |
| `deploy-render.yml` | CI exitoso en `main` | Espera el commit en `/salud` de producción y hace smoke test |
| `deploy.yml` | CI exitoso en `main` | Publica el tablero estático en GitHub Pages (heredado de la EA2) |

Render está configurado con `render.yaml` y `autoDeployTrigger: checksPass`: solo despliega un commit de `main`
cuando sus checks de GitHub Actions pasan.
""")

code("""
print((RAIZ / "render.yaml").read_text(encoding="utf-8"))
""")

code("""
jobs = pd.DataFrame(D["jobs_ci"]["jobs"]).rename(columns={"nombre": "Job", "duracion_s": "Duración (s)",
                                                         "conclusion": "Resultado"})
print(f"Última ejecución de CI en main ({D['jobs_ci']['sha']}): los jobs corren en paralelo")
jobs
""")

md("## 7. Evidencia: pull requests y ejecuciones de GitHub Actions")

code("""
prs = pd.DataFrame(D["pull_requests"])
prs["checks"] = prs["checks_ok"].astype(str) + "/" + prs["checks_total"].astype(str)
prs["resultado"] = prs["fusionado"].map({True: "Fusionado", False: "Cerrado sin fusionar"})
prs[["numero", "rama", "titulo", "checks", "checks_fallidos", "resultado"]]
""")

code("""
ejec = pd.DataFrame(D["ejecuciones"])
(ejec.groupby(["workflow", "conclusion"])["duracion_s"]
     .agg(ejecuciones="count", mediana_s="median", max_s="max").reset_index())
""")

code("""
import matplotlib.pyplot as plt
from matplotlib.patches import Patch

colores = {"CI": "#2a78d6", "Deploy": "#8a919c", "Deploy Render": "#0e7c7b"}
fig, ax = plt.subplots(figsize=(9, 3.2))
color = [("#c0392b" if c == "failure" else colores[w])
         for w, c in zip(ejec["workflow"], ejec["conclusion"], strict=True)]
ax.bar(range(1, len(ejec) + 1), ejec["duracion_s"], color=color)
ax.set(xlabel="Ejecución (orden cronológico)", ylabel="Segundos", title="Duración de las ejecuciones de GitHub Actions")
for s in ("top", "right"):
    ax.spines[s].set_visible(False)
ax.legend(handles=[Patch(color=v, label=k) for k, v in colores.items()] + [Patch(color="#c0392b", label="Fallo")],
          frameon=False, ncol=4, loc="upper left")
plt.show()
""")

md("""
## 8. La app en producción (Render)

Estas celdas consultan la app desplegada. El plan gratuito de Render apaga el servicio cuando no hay tráfico, así
que la primera petición puede tardar unos 50 segundos.
""")

code("""
def get(ruta):
    with urllib.request.urlopen("https://ea3-galaxy-a55.onrender.com" + ruta, timeout=120) as r:
        return r.status, json.loads(r.read())

estado, salud = get("/salud")
main_sha = subprocess.run(["git", "rev-parse", "main"], cwd=RAIZ, capture_output=True, text=True).stdout.strip()
print("HTTP", estado, salud)
print("¿Producción sirve el commit de main?", salud["commit"] == main_sha, f"({main_sha[:7]})")
""")

code("""
_, resumen_prod = get("/api/resumen")
print("¿El resumen en producción es idéntico al del repositorio?", resumen_prod == resumen)
pd.Series(D["produccion"]["latencia_ms"], name="Latencia mediana (ms)").rename_axis("Ruta").to_frame()
""")

code("""
pd.DataFrame(D["eventos_render"]).rename(columns={"commit": "Commit", "evento": "Evento en Render",
                                                  "inicio": "Inicio", "en_linea": "En línea"})
""")

md("""
## 9. Prueba de fallo controlada

En el PR #11 saqué a propósito `a55_limpio.csv` de la imagen Docker. pytest siguió pasando porque usa el
repositorio completo, pero el smoke test del job `docker` recibió un **503** en `/api/publicaciones`. El check `ci`
quedó en rojo, GitHub bloqueó el merge y Render no desplegó nada.
""")

code("""
demo = next(p for p in D["pull_requests"] if p["numero"] == 11)
pd.Series({"PR": f"#{demo['numero']} {demo['titulo']}",
           "Checks en verde": f"{demo['checks_ok']}/{demo['checks_total']}",
           "Checks fallidos": ", ".join(demo["checks_fallidos"]),
           "Fusionado": demo["fusionado"]}, name="").to_frame()
""")

md("""
## 10. Conclusiones

- La app web reutiliza el código de las evidencias anteriores en lugar de reescribirlo: el mismo `config.py` y el
  mismo tablero sirven para GitHub Pages y para la app en Render.
- Cada cambio pasa por el mismo camino: issue, rama, pull request, CI en verde, merge y despliegue automático.
  Nada llega a producción sin pruebas en verde, y el workflow de despliegue comprueba que producción sirva
  exactamente el commit validado.
- Probar la **imagen** y no solo el código paga: el fallo controlado del PR #11 pasó todas las pruebas de pytest y
  solo lo detectó el smoke test del contenedor.
- Usar `checksPass` en Render dejó el despliegue continuo sin tokens guardados en GitHub.
""")

nb = nbf.v4.new_notebook()
nb["cells"] = celdas
nb["metadata"]["kernelspec"] = {"name": "python3", "display_name": "Python 3", "language": "python"}
nbf.write(nb, DESTINO)
print(f"Notebook escrito en {DESTINO}")
