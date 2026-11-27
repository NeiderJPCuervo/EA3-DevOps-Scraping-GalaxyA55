"""Recolecta los datos reales del repositorio y de producción para el informe de la EA3.

Usa la CLI de GitHub (gh, con sesión iniciada), git y Docker, ejecuta pytest con
cobertura y consulta la app desplegada en Render. Escribe informe/datos_ea3.json,
que leen generar_informe_ea3.py y el notebook.

Uso (desde la raíz del repositorio):
    python informe/recolectar_datos_ea3.py
"""
import json
import os
import re
import subprocess
import sys
import tempfile
import time
import urllib.request
from collections import Counter
from datetime import datetime
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
REPO = "NeiderJPCuervo/EA3-DevOps-Scraping-GalaxyA55"
APP = "https://ea3-galaxy-a55.onrender.com"
INICIO_EA3 = "2026-11-27T00:00:00"     # el repositorio de la EA3 se creó ese día a partir del de la EA2
SALIDA = RAIZ / "informe" / "datos_ea3.json"

# Eventos del panel de Render (Events del servicio ea3-galaxy-a55), copiados del panel;
# la API de Render requiere una clave que este proyecto no guarda.
EVENTOS_RENDER = [
    {"commit": "e00944c", "evento": "Primer despliegue (creación del Blueprint)",
     "inicio": "09:52", "en_linea": "09:53"},
    {"commit": "5392895", "evento": "Despliegue automático tras el CI (New commit via Auto-Deploy)",
     "inicio": "09:56", "en_linea": "09:57"},
]


def sh(*cmd):
    return subprocess.run(cmd, cwd=RAIZ, check=True, capture_output=True, text=True).stdout


def api(ruta):
    return json.loads(sh("gh", "api", ruta))


def segundos(inicio, fin):
    a = datetime.fromisoformat(inicio.replace("Z", "+00:00"))
    b = datetime.fromisoformat(fin.replace("Z", "+00:00"))
    return round((b - a).total_seconds())


def commits():
    formato = "%h%x1f%ad%x1f%s%x1f%P"
    filas = []
    for linea in sh("git", "log", "main", f"--since={INICIO_EA3}", "--date=iso-strict",
                    f"--pretty=format:{formato}").splitlines():
        h, fecha, asunto, padres = linea.split("\x1f")
        filas.append({"hash": h, "fecha": fecha, "asunto": asunto, "merge": len(padres.split()) > 1})
    return filas[::-1]


def pull_requests():
    prs = []
    for p in sorted(api(f"repos/{REPO}/pulls?state=all&per_page=100"), key=lambda x: x["number"]):
        runs = api(f"repos/{REPO}/commits/{p['head']['sha']}/check-runs?per_page=100")["check_runs"]
        prs.append({
            "numero": p["number"], "titulo": p["title"], "rama": p["head"]["ref"],
            "fusionado": bool(p["merged_at"]), "estado": p["state"],
            "checks_ok": sum(c["conclusion"] == "success" for c in runs),
            "checks_total": len(runs),
            "checks_fallidos": sorted(c["name"] for c in runs if c["conclusion"] == "failure"),
            "issues": [int(n) for n in re.findall(r"(?:Closes|Refs) #(\d+)", p["body"] or "")],
        })
    return prs


def issues():
    return [{"numero": i["number"], "titulo": i["title"], "estado": i["state"]}
            for i in api(f"repos/{REPO}/issues?state=all&per_page=100") if "pull_request" not in i]


def ejecuciones():
    runs = api(f"repos/{REPO}/actions/runs?per_page=100")["workflow_runs"]
    return [{
        "id": r["id"], "workflow": r["name"], "evento": r["event"], "rama": r["head_branch"],
        "sha": r["head_sha"][:7], "conclusion": r["conclusion"], "titulo": r["display_title"],
        "inicio": r["run_started_at"], "fin": r["updated_at"],
        "duracion_s": segundos(r["run_started_at"], r["updated_at"]),
    } for r in sorted(runs, key=lambda r: r["run_started_at"])]


def jobs_ci():
    """Duración de cada job en la última ejecución de CI sobre main."""
    run = next(r for r in api(f"repos/{REPO}/actions/runs?branch=main&per_page=50")["workflow_runs"]
               if r["name"] == "CI" and r["conclusion"] == "success")
    jobs = api(f"repos/{REPO}/actions/runs/{run['id']}/jobs")["jobs"]
    return {"run_id": run["id"], "sha": run["head_sha"][:7],
            "jobs": [{"nombre": j["name"], "duracion_s": segundos(j["started_at"], j["completed_at"]),
                      "conclusion": j["conclusion"]} for j in jobs]}


def cobertura():
    with tempfile.TemporaryDirectory() as tmp:
        reporte = Path(tmp) / "cov.json"
        res = subprocess.run([sys.executable, "-m", "pytest", "-q", "--cov", f"--cov-report=json:{reporte}"],
                             cwd=RAIZ, capture_output=True, text=True, env=os.environ | {"EA2_SALIDA": tmp})
        datos = json.loads(reporte.read_text())
    archivos = {Path(k).as_posix(): round(v["summary"]["percent_covered"], 1) for k, v in datos["files"].items()}
    return {"resultado": res.stdout.strip().splitlines()[-1], "total": round(datos["totals"]["percent_covered"], 1),
            "archivos": archivos}


def pruebas_por_archivo():
    salida = sh(sys.executable, "-m", "pytest", "--collect-only", "-q")
    return dict(Counter(linea.split("::")[0] for linea in salida.splitlines() if "::" in linea))


def imagen_docker():
    sh("docker", "build", "-q", "-t", "ea3-galaxy-a55:informe", ".")
    tam = int(sh("docker", "image", "inspect", "ea3-galaxy-a55:informe", "--format", "{{.Size}}").strip())
    return {"tamano_mb": round(tam / 1e6), "base": "python:3.12-slim"}


def produccion():
    def get(ruta):
        inicio = time.perf_counter()
        with urllib.request.urlopen(APP + ruta, timeout=90) as r:
            cuerpo = r.read()
            return r.status, (time.perf_counter() - inicio) * 1000, cuerpo

    get("/salud")                                   # despierta el servicio del plan free
    estado, _, cuerpo = get("/salud")
    salud = json.loads(cuerpo)
    latencias = {}
    for ruta in ("/salud", "/api/resumen", "/api/publicaciones?condicion=Nuevo", "/"):
        tiempos = [get(ruta)[1] for _ in range(5)]
        latencias[ruta] = round(sorted(tiempos)[2])   # mediana de 5
    publicaciones = json.loads(get("/api/publicaciones?condicion=Nuevo")[2])
    return {"url": APP, "estado_http": estado, "salud": salud, "latencia_ms": latencias,
            "nuevos": publicaciones["total"], "consultado": datetime.now().isoformat(timespec="seconds")}


def main():
    prot = api(f"repos/{REPO}/branches/main/protection")
    datos = {
        "repo": REPO,
        "url_repo": f"https://github.com/{REPO}",
        "url_pages": api(f"repos/{REPO}/pages")["html_url"],
        "fecha": datetime.now().isoformat(timespec="seconds"),
        "main": sh("git", "rev-parse", "--short", "main").strip(),
        "commits": commits(),
        "pull_requests": pull_requests(),
        "issues": issues(),
        "ejecuciones": ejecuciones(),
        "jobs_ci": jobs_ci(),
        "cobertura": cobertura(),
        "pruebas_por_archivo": pruebas_por_archivo(),
        "imagen": imagen_docker(),
        "produccion": produccion(),
        "eventos_render": EVENTOS_RENDER,
        "proteccion_main": {
            "checks": prot["required_status_checks"]["contexts"],
            "rama_actualizada": prot["required_status_checks"]["strict"],
            "aplica_a_admins": prot["enforce_admins"]["enabled"],
        },
    }
    SALIDA.write_text(json.dumps(datos, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"{len(datos['commits'])} commits, {len(datos['pull_requests'])} PR, {len(datos['issues'])} issues, "
          f"{len(datos['ejecuciones'])} ejecuciones, cobertura {datos['cobertura']['total']} %, "
          f"producción en {datos['produccion']['salud']['commit'][:7]} → {SALIDA}")


if __name__ == "__main__":
    main()
