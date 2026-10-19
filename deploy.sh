#!/usr/bin/env bash
# ============================================================
#  Размещение веб-сервиса «Билетоматы BPA/APB» на VPS
#  http://78.111.91.152:9000
#
#  Запускать с машины, у которой есть SSH-доступ к VPS:
#     ./deploy.sh                 # ключ по умолчанию (~/.ssh/id_rsa)
#     ./deploy.sh /path/key.pem   # свой приватный ключ
#  Либо задайте пароль через переменную (нужен sshpass):
#     VPS_PASS='ваш_пароль' ./deploy.sh
# ============================================================
set -euo pipefail

VPS_HOST="${VPS_HOST:-78.111.91.152}"
VPS_USER="${VPS_USER:-root}"
VPS_PORT_SSH="${VPS_PORT_SSH:-22}"
APP_DIR="/opt/cppk"
PORT=9000

KEY="${1:-$HOME/.ssh/id_rsa}"
if [[ -n "${VPS_PASS:-}" ]] && command -v sshpass >/dev/null; then
  SSHP=(sshpass -p "$VPS_PASS")
  SCPX=(sshpass -p "$VPS_PASS" scp -P "$VPS_PORT_SSH" -o StrictHostKeyChecking=no)
  SSHX=(sshpass -p "$VPS_PASS" ssh -p "$VPS_PORT_SSH" -o StrictHostKeyChecking=no)
else
  SSHP=()
  SCPX=(scp -i "$KEY" -P "$VPS_PORT_SSH" -o StrictHostKeyChecking=no)
  SSHX=(ssh -i "$KEY" -p "$VPS_PORT_SSH" -o StrictHostKeyChecking=no)
fi

cd "$(dirname "$0")"

echo "==> 1. Копируем файлы приложения на $VPS_USER@$VPS_HOST ..."
"${SSHP[@]}" ssh -p "$VPS_PORT_SSH" -o StrictHostKeyChecking=no \
  "$VPS_USER@$VPS_HOST" "mkdir -p $APP_DIR"

"${SCPX[@]}" app.py data.py xls_import.py page_ui.html requirements.txt cppk.service \
  "$VPS_USER@$VPS_HOST:$APP_DIR/"

# если есть уже наполненная база объектов — берём её с собой
[[ -f upload_store.json ]] && "${SCPX[@]}" upload_store.json "$VPS_USER@$VPS_HOST:$APP_DIR/" || true

echo "==> 2. Установка зависимостей и systemd-сервиса на VPS ..."
"${SSHP[@]}" ssh -p "$VPS_PORT_SSH" -o StrictHostKeyChecking=no "$VPS_USER@$VPS_HOST" "bash -s" <<'REMOTE'
set -e
cd /opt/cppk

# python3 + pip
if ! command -v python3 >/dev/null; then
  (apt-get update -y && apt-get install -y python3 python3-pip) \
  || (yum install -y python3 python3-pip) \
  || (dnf install -y python3 python3-pip)
fi
python3 -m pip --version >/dev/null 2>&1 || \
  (apt-get update -y && apt-get install -y python3-pip) || \
  (yum install -y python3-pip) || (dnf install -y python3-pip)

python3 -m pip install --quiet --upgrade fastapi "uvicorn[standard]" openpyxl xlrd

# systemd unit (ExecStart=/usr/bin/python3 — при необходимости поправьте путь)
cp -f cppk.service /etc/systemd/system/cppk.service
sed -i "s|ExecStart=.*|ExecStart=$(command -v python3) /opt/cppk/app.py|" /etc/systemd/system/cppk.service
systemctl daemon-reload
systemctl enable cppk
systemctl restart cppk

# firewall — открыть 9000, если фаервол активен
if command -v ufw >/dev/null && ufw status | grep -q Active; then ufw allow 9000/tcp; fi
if command -v firewall-cmd >/dev/null && firewall-cmd --state 2>/dev/null | grep -q running; then
  firewall-cmd --permanent --add-port=9000/tcp && firewall-cmd --reload
fi
# iptables без frontend-утилит
if command -v iptables >/dev/null; then iptables -I INPUT -p tcp --dport 9000 -j ACCEPT 2>/dev/null || true; fi

sleep 2
systemctl status cppk --no-pager | head -6
curl -s http://localhost:9000/api/health && echo "  <- health OK"
REMOTE

echo ""
echo "==> Готово! Сервис доступен по адресу: http://${VPS_HOST}:${PORT}/"
echo "    Логи на VPS: journalctl -u cppk -f"
echo "    Обновление кода: повторный запуск ./deploy.sh"
