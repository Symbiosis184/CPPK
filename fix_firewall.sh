#!/usr/bin/env bash
# Диагностика и открытие порта 9000 на VPS. Запуск от root:
#   curl -fsSL https://raw.githubusercontent.com/Symbiosis184/CPPK/main/fix_firewall.sh | bash
set -uo pipefail
PORT="${PORT:-9000}"

echo "== 1. Статус службы cppk =="
systemctl status cppk --no-pager -l | head -n 12 || echo "служба cppk НЕ установлена"
echo "-- последние логи сервиса --"
journalctl -u cppk -n 15 --no-pager 2>/dev/null || true

echo
echo "== 2. Кто слушает порт $PORT =="
( ss -tlnp 2>/dev/null || netstat -tlnp ) | grep -E ":$PORT\b" || {
  echo "ПОРТ $PORT НИКТО НЕ СЛУШАЕТ! Сервис не запущен."
  echo "Установите его командой:"
  echo "  curl -fsSL https://raw.githubusercontent.com/Symbiosis184/CPPK/main/setup_from_git.sh | bash"
}

echo
echo "== 3. Проверка локально (127.0.0.1) =="
if curl -sf -m 5 "http://127.0.0.1:$PORT/api/health"; then
  echo; echo "Локально сервис ОТВЕЧАЕТ -> проблема в фаерволе, идем дальше"
  LOCAL_OK=1
else
  echo "Локально сервис НЕ отвечает -> проблема не в фаерволе, а в самом сервисе (см. статус выше)"
  LOCAL_OK=0
fi

echo
echo "== 4. Разрешаем порт в локальных фаерволах =="
CHANGED=0
if command -v ufw >/dev/null 2>&1 && ufw status 2>/dev/null | grep -q "Status: active"; then
  ufw allow "$PORT/tcp" comment 'CPPK' >/dev/null && echo "ufw: порт $PORT открыт"; CHANGED=1
else echo "ufw: неактивен/не установлен — пропускаю"; fi

if command -v firewall-cmd >/dev/null 2>&1 && firewall-cmd --state >/dev/null 2>&1; then
  firewall-cmd --permanent --add-port="$PORT/tcp" >/dev/null && \
  firewall-cmd --reload >/dev/null && echo "firewalld: порт $PORT открыт"; CHANGED=1
else echo "firewalld: неактивен/не установлен — пропускаю"; fi

if command -v iptables >/dev/null 2>&1; then
  if iptables -L INPUT -n 2>/dev/null | grep -qE "REJECT|DROP"; then
    iptables -C INPUT -p tcp --dport "$PORT" -j ACCEPT 2>/dev/null || {
      iptables -I INPUT -p tcp --dport "$PORT" -j ACCEPT && echo "iptables: правило добавлено"; CHANGED=1; }
    # сохраняем правило, если есть netfilter-persistent
    command -v netfilter-persistent >/dev/null 2>&1 && netfilter-persistent save >/dev/null 2>&1 || true
  else echo "iptables: правил DROP/REJECT нет — пропускаю"; fi
fi

if [ "$CHANGED" = "0" ]; then echo "Локальные фаерволы порту не мешают."; fi

echo
echo "== 5. Итог =="
IP=$(hostname -I | awk '{print $1}')
if [ "$LOCAL_OK" = "1" ]; then
  echo "Сервис работает локально. Если http://$IP:$PORT ВСЕ ЕЩЁ недоступен снаружи,"
  echo "значит порт блокирует ВНЕШНИЙ фаервол хостинга (не на сервере)."
  echo "Действия: зайдите в панель управления вашего VPS-провайдера"
  echo "(Selectel / timeweb / reg.ru / dgis и т.п.) -> раздел 'Сетевые экраны',"
  echo "'Firewall', 'Безопасность' или 'Группы безопасности' -> создайте правило:"
  echo "  Inbound (входящие), протокол TCP, порт $PORT, источник 0.0.0.0/0, действие ALLOW."
  echo "Если такой панели нет — напишите в поддержку хостинга: 'разрешите входящий TCP порт $PORT'."
fi
echo "FIX_FIREWALL_DONE"
