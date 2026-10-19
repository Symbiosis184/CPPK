#!/usr/bin/env bash
# Установка/обновление CPPK-сервиса с GitHub. Запуск на VPS от root:
#   curl -fsSL https://raw.githubusercontent.com/<ВашЛогин>/<Репозиторий>/main/setup_from_git.sh | bash
set -euo pipefail

REPO_URL="${REPO_URL:-https://github.com/USER/REPO.git}"  # можно переопределить: REPO_URL=... bash setup_from_git.sh
APP_DIR="/opt/cppk"
PORT="${PORT:-9000}"

if [ "$(id -u)" != "0" ]; then echo "Запустите от root"; exit 1; fi
if ! command -v git >/dev/null; then
  apt-get update -qq && apt-get install -y -qq git ca-certificates
fi

# 1. Клонируем или обновляем репозиторий
if [ -d "$APP_DIR/.git" ]; then
  echo "== Обновление из git =="
  git -C "$APP_DIR" fetch --all -q
  BRANCH="$(git -C "$APP_DIR" rev-parse --abbrev-ref HEAD)"
  git -C "$APP_DIR" reset --hard -q "origin/$BRANCH"
else
  echo "== Клонирование репозитория =="
  if [ -d "$APP_DIR" ] && [ -n "$(ls -A "$APP_DIR" 2>/dev/null)" ]; then
    echo "Каталог $APP_DIR существует и не пуст, но это не git-репозиторий."
    echo "Укажите URL репозитория:  REPO_URL=https://github.com/USER/REPO.git bash setup_from_git.sh"
    exit 1
  fi
  mkdir -p "$APP_DIR"
  git clone -q "$REPO_URL" "$APP_DIR"
fi

# 2. Зависимости Python в venv (обход PEP 668 на Debian 13)
apt-get update -qq
apt-get install -y -qq python3-venv python3-pip libjemalloc2
if [ ! -x "$APP_DIR/venv/bin/python" ]; then
  python3 -m venv "$APP_DIR/venv"
fi
"$APP_DIR/venv/bin/pip" install -q --upgrade pip
"$APP_DIR/venv/bin/pip" install -q -r "$APP_DIR/requirements.txt"

# 3. systemd-сервис
cp "$APP_DIR/cppk.service" /etc/systemd/system/cppk.service
systemctl daemon-reload
systemctl enable cppk >/dev/null 2>&1
systemctl restart cppk
sleep 2

# 4. Открыть порт в локальном фаерволе (если включён)
if command -v ufw >/dev/null && ufw status | grep -q "Status: active"; then
  ufw allow "$PORT/tcp" >/dev/null || true
fi

# 5. Проверка
STATUS=$(systemctl is-active cppk || true)
echo "== cppk: $STATUS =="
curl -sf "http://127.0.0.1:$PORT/api/health" && echo
echo "INSTALL_DONE  http://$(hostname -I | awk '{print $1}'):$PORT"
