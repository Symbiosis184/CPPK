#!/usr/bin/env bash
# Установка CPPK-сервиса на Debian 13 (запускать как root в каталоге /opt/cppk)
set -e
cd /opt/cppk

apt-get update -qq
apt-get install -y -qq python3-venv libjemalloc2

python3 -m venv venv
./venv/bin/pip install -q --upgrade pip
./venv/bin/pip install -q fastapi "uvicorn[standard]" openpyxl xlrd

cat > /etc/systemd/system/cppk.service <<'EOF'
[Unit]
Description=SberService BPA/APB network settings web service
After=network.target

[Service]
Type=simple
WorkingDirectory=/opt/cppk
ExecStart=/opt/cppk/venv/bin/python /opt/cppk/app.py
Restart=always
RestartSec=3
Environment=PORT=9000

[Install]
WantedBy=multi-user.target
EOF

systemctl daemon-reload
systemctl enable --now cppk
sleep 2
systemctl --no-pager status cppk | head -6
curl -s http://127.0.0.1:9000/api/health && echo
echo "INSTALL_DONE"
