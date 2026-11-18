# Параметры сборки для закрытого контура (все необязательные):
#   BASE_IMAGE       базовый образ из внутреннего registry, напр. harbor.corp/library/python:3.12-slim-bookworm
#   APT_MIRROR       зеркало Debian, напр. http://nexus.corp/repository/debian  (вместо deb.debian.org)
#   PIP_INDEX_URL    PyPI-прокси, напр. https://nexus.corp/repository/pypi/simple
#   PIP_TRUSTED_HOST хост PyPI-прокси, если у него внутренний сертификат
ARG BASE_IMAGE=python:3.12-slim-bookworm
FROM ${BASE_IMAGE}

ARG APT_MIRROR=""
ARG PIP_INDEX_URL=""
ARG PIP_TRUSTED_HOST=""

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1 \
    DATA_DIR=/data

# Tesseract с русским и английским (+ osd для автоповорота страниц),
# шрифты Liberation (метрически совместимы с Times New Roman / Arial) для распознанного текста.
RUN if [ -n "$APT_MIRROR" ]; then \
      sed -i "s|http://deb.debian.org/debian|${APT_MIRROR}|g" /etc/apt/sources.list.d/debian.sources 2>/dev/null || \
      sed -i "s|http://deb.debian.org/debian|${APT_MIRROR}|g" /etc/apt/sources.list; \
    fi \
 && apt-get update \
 && apt-get install -y --no-install-recommends \
      tesseract-ocr tesseract-ocr-rus tesseract-ocr-eng tesseract-ocr-osd \
      fonts-liberation fonts-dejavu-core \
      libglib2.0-0 \
 && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install ${PIP_INDEX_URL:+--index-url "$PIP_INDEX_URL"} \
                ${PIP_TRUSTED_HOST:+--trusted-host "$PIP_TRUSTED_HOST"} \
                -r requirements.txt

COPY app ./app
COPY static ./static

RUN useradd --create-home --uid 10001 appuser \
 && mkdir -p /data && chown appuser:appuser /data
USER 10001
VOLUME ["/data"]
EXPOSE 8000

HEALTHCHECK --interval=30s --timeout=5s --start-period=20s --retries=3 \
  CMD python -c "import urllib.request,sys; sys.exit(0 if urllib.request.urlopen('http://127.0.0.1:8000/healthz', timeout=4).status == 200 else 1)"

# Один процесс uvicorn: параллельность обеспечивают рабочие процессы конвертации (CONVERT_WORKERS)
CMD ["uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8000", "--proxy-headers", "--forwarded-allow-ips", "*"]
