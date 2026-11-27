# Imagen de la app web de la EA3 (FastAPI + uvicorn). Solo lleva lo que la app
# necesita en producción: los datos ya procesados, las figuras y dos módulos del
# scraper (config.py y 06_sitio.py). Selenium, Scrapy y el HTML capturado no viajan.
FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    PORT=8000 \
    APP_ENV=produccion

WORKDIR /srv

COPY requirements-app.txt .
RUN pip install -r requirements-app.txt

# Usuario sin privilegios: la app solo necesita leer los datos
RUN useradd --create-home --uid 10001 app
COPY app/ app/
COPY scraping/config.py scraping/06_sitio.py scraping/
COPY data/resumen.json data/a55_limpio.csv data/metricas_selenium.json data/
COPY figuras/ figuras/
# config.py crea data/raw al importarse: se deja lista y con dueño "app"
RUN mkdir -p data/raw && chown -R app:app /srv
USER app

EXPOSE 8000
HEALTHCHECK --interval=30s --timeout=5s --start-period=10s --retries=3 \
    CMD python -c "import os, urllib.request; urllib.request.urlopen(f'http://127.0.0.1:{os.environ[\"PORT\"]}/salud', timeout=4)"

# Render asigna el puerto en $PORT
CMD ["sh", "-c", "exec uvicorn app.main:app --host 0.0.0.0 --port ${PORT} --proxy-headers --forwarded-allow-ips='*'"]
