"""App web de la EA3: tablero y API del scraper del Samsung Galaxy A55.

Se despliega en Render como contenedor Docker. El puerto lo define la variable
PORT (Render la asigna); el commit y el entorno salen de variables de entorno
para que el CD pueda verificar qué versión quedó desplegada.
"""
import importlib.util
import json
import os
import sys
from pathlib import Path

import pandas as pd
from fastapi import FastAPI, HTTPException, Query
from fastapi.responses import HTMLResponse
from fastapi.staticfiles import StaticFiles

# --- Reutilización del código de la EA1/EA2 ------------------------------------
# config.py define DATA_DIR y FIG_DIR (con la variable EA2_SALIDA para CI y pruebas)
# y 06_sitio.py construye el mismo tablero que se publica en GitHub Pages.
SCRAPING = Path(__file__).resolve().parent.parent / "scraping"
sys.path.insert(0, str(SCRAPING))
import config  # noqa: E402

_spec = importlib.util.spec_from_file_location("sitio", SCRAPING / "06_sitio.py")
sitio = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(sitio)

COLUMNAS_API = ["id", "titulo", "condicion", "capacidad_gb", "precio_actual", "precio_anterior",
                "descuento_pct", "envio_gratis", "rating", "url"]

app = FastAPI(
    title="Galaxy A55 · Mercado Libre Colombia",
    description="Tablero y API del análisis de precios del Samsung Galaxy A55 (EA3).",
    version="1.0.0",
)
app.mount("/figuras", StaticFiles(directory=config.FIG_DIR), name="figuras")


def commit_desplegado() -> str:
    # Render define RENDER_GIT_COMMIT con el commit desplegado; en local no existe.
    return os.environ.get("RENDER_GIT_COMMIT", "local")


def leer_resumen() -> dict:
    ruta = config.DATA_DIR / "resumen.json"
    if not ruta.exists():
        raise HTTPException(status_code=503, detail="Datos no disponibles: falta resumen.json. "
                                                    "Ejecute 04_analisis.py para generarlo.")
    return json.loads(ruta.read_text(encoding="utf-8"))


def leer_publicaciones() -> pd.DataFrame:
    ruta = config.DATA_DIR / "a55_limpio.csv"
    if not ruta.exists():
        raise HTTPException(status_code=503, detail="Datos no disponibles: falta a55_limpio.csv. "
                                                    "Ejecute 04_analisis.py para generarlo.")
    df = pd.read_csv(ruta)
    df["url"] = df["url"].str.split("#").str[0]          # sin parámetros de rastreo
    df["capacidad_gb"] = df["capacidad_gb"].astype("Int64")
    return df


def a_registros(df: pd.DataFrame) -> list[dict]:
    """Filas como diccionarios JSON: los NaN de pandas se convierten en null."""
    return df[COLUMNAS_API].astype(object).where(df[COLUMNAS_API].notna(), None).to_dict(orient="records")


@app.get("/", response_class=HTMLResponse, summary="Tablero de precios")
def tablero() -> str:
    """Tablero HTML con indicadores, gráficas y publicaciones del Galaxy A55."""
    commit = commit_desplegado()
    origen = (f'commit <a href="https://github.com/{sitio.REPO}/commit/{commit}"><code>{commit[:7]}</code></a> '
              "desplegado en Render" if commit != "local" else "servidor local")
    return sitio.construir(leer_resumen(), leer_publicaciones(), origen=origen)


@app.get("/salud", summary="Chequeo de salud")
def salud() -> dict:
    """Chequeo de salud: lo usan Render (health check) y el CD (verificación del despliegue)."""
    return {
        "estado": "ok",
        "commit": commit_desplegado(),
        "entorno": os.environ.get("APP_ENV", "desarrollo"),
    }


@app.get("/api/resumen", summary="Resumen del análisis")
def resumen() -> dict:
    """Indicadores del análisis: totales, estadísticos de precio, condición, envío y descuentos."""
    return leer_resumen()


@app.get("/api/publicaciones", summary="Publicaciones del Galaxy A55")
def publicaciones(
    condicion: str | None = Query(None, description="Nuevo, Usado o Caja abierta", examples=["Usado"]),
    capacidad: int | None = Query(None, description="Capacidad en GB", examples=[256]),
    precio_max: int | None = Query(None, ge=0, description="Precio máximo en COP", examples=[1_000_000]),
    limite: int = Query(50, ge=1, le=200, description="Máximo de publicaciones a devolver"),
) -> dict:
    """Publicaciones del Galaxy A55 ordenadas por precio, con filtros opcionales."""
    df = leer_publicaciones()
    if condicion:
        df = df[df["condicion"].str.casefold() == condicion.casefold()]
    if capacidad is not None:
        df = df[df["capacidad_gb"] == capacidad]
    if precio_max is not None:
        df = df[df["precio_actual"] <= precio_max]
    df = df.sort_values("precio_actual")
    return {"total": len(df), "publicaciones": a_registros(df.head(limite))}


@app.get("/api/publicaciones/{id_publicacion}", summary="Una publicación")
def publicacion(id_publicacion: str) -> dict:
    """Detalle de una publicación por su identificador de Mercado Libre (MCO...)."""
    df = leer_publicaciones()
    fila = df[df["id"] == id_publicacion.upper()]
    if fila.empty:
        raise HTTPException(status_code=404, detail=f"No existe la publicación {id_publicacion}.")
    return a_registros(fila)[0]
