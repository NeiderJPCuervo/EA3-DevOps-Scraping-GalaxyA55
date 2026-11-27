"""Pruebas de la app web FastAPI (EA3): tablero, salud y API."""
import pytest
from fastapi.testclient import TestClient

from conftest import RAIZ


@pytest.fixture
def cliente(monkeypatch):
    from app import main
    # La app lee los datos versionados de la EA1 (en las pruebas config apunta a una carpeta temporal)
    monkeypatch.setattr(main.config, "DATA_DIR", RAIZ / "data")
    monkeypatch.delenv("RENDER_GIT_COMMIT", raising=False)
    monkeypatch.delenv("APP_ENV", raising=False)
    return TestClient(main.app)


def test_salud_en_local(cliente):
    r = cliente.get("/salud")
    assert r.status_code == 200
    assert r.json() == {"estado": "ok", "commit": "local", "entorno": "desarrollo"}


def test_salud_reporta_el_commit_desplegado(cliente, monkeypatch):
    monkeypatch.setenv("RENDER_GIT_COMMIT", "abc1234def5678")
    monkeypatch.setenv("APP_ENV", "produccion")
    assert cliente.get("/salud").json() == {"estado": "ok", "commit": "abc1234def5678", "entorno": "produccion"}


def test_resumen(cliente):
    r = cliente.get("/api/resumen")
    assert r.status_code == 200
    datos = r.json()
    assert datos["publicaciones_totales"] == 119
    assert datos["publicaciones_a55"] == 30


def test_sin_datos_responde_503(cliente, monkeypatch, tmp_path):
    from app import main
    monkeypatch.setattr(main.config, "DATA_DIR", tmp_path)
    for ruta in ("/api/resumen", "/api/publicaciones", "/"):
        r = cliente.get(ruta)
        assert r.status_code == 503
        assert "Datos no disponibles" in r.json()["detail"]


def test_publicaciones_ordenadas_y_limpias(cliente):
    datos = cliente.get("/api/publicaciones").json()
    assert datos["total"] == 30
    precios = [p["precio_actual"] for p in datos["publicaciones"]]
    assert precios == sorted(precios)
    assert all("#" not in p["url"] for p in datos["publicaciones"])
    assert all(p["id"].startswith("MCO") for p in datos["publicaciones"])


@pytest.mark.parametrize("parametros, condicion", [
    ({"condicion": "usado"}, lambda p: p["condicion"] == "Usado"),
    ({"condicion": "Nuevo"}, lambda p: p["condicion"] == "Nuevo"),
    ({"capacidad": 256}, lambda p: p["capacidad_gb"] == 256),
    ({"precio_max": 800_000}, lambda p: p["precio_actual"] <= 800_000),
])
def test_filtros(cliente, parametros, condicion):
    datos = cliente.get("/api/publicaciones", params=parametros).json()
    assert datos["total"] > 0
    assert all(condicion(p) for p in datos["publicaciones"])


def test_filtros_combinados_y_limite(cliente):
    datos = cliente.get("/api/publicaciones", params={"condicion": "Usado", "capacidad": 256, "limite": 5}).json()
    assert datos["total"] == 19
    assert len(datos["publicaciones"]) == 5


@pytest.mark.parametrize("parametros", [{"precio_max": -1}, {"limite": 0}, {"limite": 500}, {"capacidad": "mucha"}])
def test_parametros_invalidos_responden_422(cliente, parametros):
    assert cliente.get("/api/publicaciones", params=parametros).status_code == 422


def test_publicacion_por_id(cliente):
    primera = cliente.get("/api/publicaciones", params={"limite": 1}).json()["publicaciones"][0]
    r = cliente.get(f"/api/publicaciones/{primera['id'].lower()}")
    assert r.status_code == 200
    assert r.json() == primera


def test_publicacion_inexistente_responde_404(cliente):
    r = cliente.get("/api/publicaciones/MCO000")
    assert r.status_code == 404
    assert "MCO000" in r.json()["detail"]


def test_tablero_html(cliente):
    r = cliente.get("/")
    assert r.status_code == 200
    assert r.headers["content-type"].startswith("text/html")
    assert "Publicaciones del Galaxy A55 (30)" in r.text
    assert "servidor local" in r.text


def test_tablero_muestra_el_commit_de_render(cliente, monkeypatch):
    monkeypatch.setenv("RENDER_GIT_COMMIT", "abc1234def5678")
    r = cliente.get("/")
    assert "<code>abc1234</code></a> desplegado en Render" in r.text


def test_documentacion_openapi(cliente):
    esquema = cliente.get("/openapi.json").json()
    assert {"/", "/salud", "/api/resumen", "/api/publicaciones"} <= set(esquema["paths"])
