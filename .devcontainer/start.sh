#!/usr/bin/env bash
# Запуск сервиса в GitHub Codespaces и публикация ссылки для тестирования.
# Запускается автоматически при открытии codespace (postAttachCommand).
set -u
cd "$(dirname "$0")/.."

PORT=8000
export DEMO_NOTICE="${DEMO_NOTICE:-Тестовый стенд. Не загружайте конфиденциальные документы — файлы удаляются с сервера через 60 минут.}"

if curl -sf "http://127.0.0.1:$PORT/healthz" >/dev/null 2>&1; then
  echo "Сервер уже запущен на порту $PORT."
  exit 0
fi

uvicorn app.main:app --host 0.0.0.0 --port "$PORT" &
SERVER_PID=$!

for _ in $(seq 60); do
  curl -sf "http://127.0.0.1:$PORT/healthz" >/dev/null 2>&1 && break
  sleep 1
done

URL="http://127.0.0.1:$PORT"
PUBLIC=0
if [ -n "${CODESPACE_NAME:-}" ]; then
  URL="https://${CODESPACE_NAME}-${PORT}.${GITHUB_CODESPACES_PORT_FORWARDING_DOMAIN:-app.github.dev}"
  # порт появляется в списке переадресованных не сразу — несколько попыток
  for _ in $(seq 10); do
    if gh codespace ports visibility "$PORT:public" -c "$CODESPACE_NAME" >/dev/null 2>&1; then
      PUBLIC=1
      break
    fi
    sleep 3
  done
fi

echo
echo "================================================================================"
echo "  PDF → Word запущен"
echo
echo "  Ссылка: $URL"
echo
if [ "$PUBLIC" = 1 ]; then
  echo "  Порт открыт публично — ссылку можно отправлять коллегам."
elif [ -n "${CODESPACE_NAME:-}" ]; then
  echo "  Не удалось открыть порт автоматически. Сделайте вручную:"
  echo "  вкладка PORTS → строка 8000 → правый клик → Port Visibility → Public."
fi
echo
echo "  Ссылка работает, пока запущен этот codespace. Он останавливается после периода"
echo "  бездействия (по умолчанию 30 мин; до 240 мин — github.com/settings/codespaces)."
echo "  Запросы тестировщиков пишутся в этот терминал и продлевают работу."
echo "================================================================================"
echo

wait "$SERVER_PID"
