# Перенос в корпоративный контур

Инструкция для развёртывания в закрытой сети без доступа в интернет (или с доступом только
через внутренние зеркала: Nexus, Harbor, Artifactory, корпоративный GitLab).

## 1. Что потребуется

| Ресурс | Минимум | Рекомендуется |
|---|---|---|
| Сервер | Linux x86_64, 2 vCPU, 4 ГБ RAM | 4 vCPU, 8 ГБ RAM |
| Диск | 5 ГБ под образ + место под файлы | 20 ГБ |
| ПО | Docker / Podman **или** Python 3.10–3.12 + Tesseract | Docker + nginx |
| Сеть | DNS-имя, TLS-сертификат | + доменная авторизация (SSO) |

Ориентир по нагрузке: одна одновременная конвертация ≈ 1 ядро и 0,5–1,5 ГБ памяти;
текстовая страница — доли секунды, скан с OCR — 2–4 с на страницу.

## 2. Сведения для службы ИБ

- **Нет исходящих соединений.** Приложение не обращается во внешнюю сеть: нет CDN, внешних
  шрифтов, телеметрии и облачных API. OCR (Tesseract) работает локально.
- **Хранение данных.** Загруженные PDF и готовые DOCX лежат только в каталоге `DATA_DIR`
  и удаляются автоматически через `FILE_TTL_MIN` минут (по умолчанию 60) или кнопкой «×».
  Базы данных нет.
- **Пароли PDF** используются только для расшифровки в памяти и не сохраняются ни на диск, ни в логи.
- **Изоляция.** Каждый файл обрабатывается в отдельном процессе с таймаутом
  (`CONVERT_TIMEOUT`) и, при необходимости, лимитом памяти (`WORKER_MEMORY_MB`);
  в контейнере — от непривилегированного пользователя (uid 10001).
- **Проверка входных данных.** Принимаются только PDF (проверка сигнатуры), лимиты на размер
  и число страниц, идентификаторы задач — случайные UUID без доступа к путям.
- **Авторизации нет** — сервис рассчитан на работу за корпоративным SSO или в закрытом сегменте
  (см. раздел 6).

### Лицензии компонентов

| Компонент | Лицензия |
|---|---|
| PyMuPDF (MuPDF) | **AGPL-3.0** или коммерческая лицензия Artifex |
| pdf2docx | MIT |
| Tesseract OCR | Apache-2.0 |
| FastAPI, uvicorn, python-docx, Pillow, numpy, OpenCV | MIT / BSD / Apache-2.0 |
| Шрифты Liberation | SIL OFL 1.1 |

> **PyMuPDF версии 2.x не использовать.** Начиная с 2.0 она тянет обязательную зависимость
> `pymupdf-layout` с лицензией PolyForm Noncommercial (только некоммерческое использование).
> В `requirements.txt` версия ограничена: `PyMuPDF<2`. При обновлении зависимостей следите,
> чтобы в окружение не попали `pymupdf-layout` и `pymupdf4llm`:
> `pip list | grep -i -E "layout|4llm"` должен быть пустым.

AGPL-3.0 у PyMuPDF требует предоставлять исходный код приложения тем, кто пользуется им по сети.
Для внутреннего сервиса это обычно закрывается тем, что исходники лежат в корпоративном GitLab
и доступны сотрудникам. Окончательное решение — за юристами и ИБ; альтернатива — коммерческая
лицензия Artifex на PyMuPDF.

## 3. Вариант А — готовый Docker-образ (проще всего)

**Снаружи** (там, где есть интернет):

```bash
# из релиза на GitHub: Releases → pdf2word-vX.Y.Z-docker-image.tar.gz (+ SHA256SUMS.txt)
# или самостоятельно:
docker pull ghcr.io/dima36s/pdf2word:latest
docker save ghcr.io/dima36s/pdf2word:latest | gzip > pdf2word-image.tar.gz
sha256sum pdf2word-image.tar.gz > pdf2word-image.sha256
```

Перенесите архив в контур установленным у вас способом (файлообменник, проверка антивирусом,
шлюз ИБ). **Внутри**:

```bash
sha256sum -c pdf2word-image.sha256
docker load -i pdf2word-image.tar.gz
# в корпоративный registry (Harbor / Nexus):
docker tag ghcr.io/dima36s/pdf2word:latest harbor.corp.local/tools/pdf2word:1.0.0
docker push harbor.corp.local/tools/pdf2word:1.0.0
```

Запуск — `docker-compose.yml` из репозитория, заменив `build: .` на
`image: harbor.corp.local/tools/pdf2word:1.0.0`:

```bash
docker compose up -d
curl -s http://127.0.0.1:8000/healthz      # {"ok":true,...}
```

## 4. Вариант Б — сборка внутри контура из исходников

1. Перенесите репозиторий в корпоративный GitLab (снаружи):

   ```bash
   git clone --mirror https://github.com/dima36s/pdf2word.git
   cd pdf2word.git && git bundle create ../pdf2word.bundle --all
   ```
   Внутри: `git clone pdf2word.bundle pdf2word`, добавьте remote корпоративного GitLab и сделайте push.

2. Соберите образ через внутренние зеркала — все параметры необязательные:

   ```bash
   docker build \
     --build-arg BASE_IMAGE=harbor.corp.local/library/python:3.12-slim-bookworm \
     --build-arg APT_MIRROR=http://nexus.corp.local/repository/debian-bookworm \
     --build-arg PIP_INDEX_URL=https://nexus.corp.local/repository/pypi-proxy/simple \
     --build-arg PIP_TRUSTED_HOST=nexus.corp.local \
     -t harbor.corp.local/tools/pdf2word:1.0.0 .
   ```

3. Для GitLab CI достаточно одного job с `docker build` (или kaniko) и этими же аргументами.

## 5. Вариант В — без Docker (systemd)

Подходит для серверов, где контейнеры запрещены. Нужен Python 3.12 (под него собраны
офлайн-пакеты из релиза; для другой версии Python скачайте колёса заново) и glibc ≥ 2.28
(Astra Linux 1.7+, РЕД ОС 7.3+, Debian 10+, RHEL/CentOS 8+).

```bash
# 1. Tesseract с русским языком и шрифты
sudo apt install tesseract-ocr tesseract-ocr-rus tesseract-ocr-osd fonts-liberation   # Astra / Debian
# sudo dnf install tesseract tesseract-langpack-rus liberation-fonts                   # РЕД ОС / RHEL

# 2. Код и пакеты (pdf2word-vX.Y.Z-wheels-py312-linux-x86_64.tar.gz из релиза)
sudo useradd --system --home /opt/pdf2word pdf2word
# исходники: на странице релиза «Source code (tar.gz)»
sudo mkdir -p /opt/pdf2word && sudo tar xzf pdf2word-X.Y.Z.tar.gz -C /opt/pdf2word --strip-components=1
cd /opt/pdf2word
sudo python3.12 -m venv .venv
sudo tar xzf pdf2word-vX.Y.Z-wheels-py312-linux-x86_64.tar.gz
sudo .venv/bin/pip install --no-index --find-links wheels -r requirements.txt

# 3. Служба
sudo cp deploy/pdf2word.service /etc/systemd/system/
sudo systemctl daemon-reload && sudo systemctl enable --now pdf2word
curl -s http://127.0.0.1:8000/healthz
```

Если вместо офлайн-архива есть Nexus PyPI-прокси:
`pip install --index-url https://nexus.corp.local/repository/pypi-proxy/simple -r requirements.txt`.

## 6. Публикация для пользователей

- **nginx + HTTPS** — готовый пример в `deploy/nginx.conf` (лимит тела запроса 110 МБ,
  таймауты, переадресация на HTTPS).
- **Авторизация.** Встроенной нет. Варианты:
  - доменная авторизация через `oauth2-proxy` + Keycloak / ADFS / Azure AD (заготовка в `deploy/nginx.conf`);
  - Kerberos/SPNEGO на уровне nginx (модуль `spnego-http-auth`) — вход без пароля с доменных ПК;
  - на время пилота — ограничение по IP-подсетям (`allow 10.0.0.0/8; deny all;`).
- **Kubernetes / OpenShift.** Одна реплика (состояние задач — в файлах), `DATA_DIR` — на `emptyDir`
  или PVC, `/tmp` — на `emptyDir`. Масштабирование — через `CONVERT_WORKERS` и ресурсы пода.
  Если нужно несколько реплик — общий PVC (RWX) и sticky-сессии на ingress.

## 7. Настройки

Все параметры — переменные окружения, полный список в `README.md`. Для корпоративного контура
обычно меняют:

```bash
MAX_FILE_MB=200          # большие сканы
CONVERT_WORKERS=4        # по числу свободных ядер
FILE_TTL_MIN=30          # строже к хранению
WORKER_MEMORY_MB=3072    # защита от «тяжёлых» PDF
OCR_DEFAULT_LANG=rus+eng
```

Дополнительные языки распознавания: в образ/на сервер доставьте пакеты `tesseract-ocr-<язык>`
(например, `deu`, `kaz`) — они автоматически появятся в списке языков интерфейса.

## 8. Проверка после развёртывания

```bash
curl -s https://pdf2word.corp.local/healthz
curl -s https://pdf2word.corp.local/api/config | python3 -m json.tool   # ocr.available = true
```

Затем откройте сервис в браузере и сконвертируйте по одному файлу каждого типа: обычный PDF из
Word, скан и PDF с паролем. Для регрессионной проверки после обновлений в репозитории есть
`pytest` (35 тестов) и `tests/samples`.

## 9. Обновление

Новая версия → новый тег `vX.Y.Z` на GitHub → GitHub Actions соберёт и проверит образ, приложит к
релизу архив образа и офлайн-пакеты → перенос по варианту А или В. Данные пользователей
переносить не нужно: файлы временные.
