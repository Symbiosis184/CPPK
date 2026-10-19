# Развёртывание на VPS http://78.111.91.152:9000

## Вариант 1 — автоматически (со своей машины, где есть доступ к VPS)

```bash
# скопируйте папку проекта к себе, затем:
./deploy.sh                      # вход по ключу ~/.ssh/id_rsa
./deploy.sh ~/keys/vps.pem       # вход по своему ключу
VPS_USER=admin ./deploy.sh       # если пользователь не root
VPS_PASS='пароль' ./deploy.sh    # вход по паролю (нужен sshpass)
```

Скрипт: копирует код в `/opt/cppk`, ставит `fastapi uvicorn openpyxl xlrd`,
регистрирует systemd-сервис `cppk` (автозапуск + перезапуск при падении),
открывает порт 9000 в фаерволе и проверяет `http://localhost:9000/api/health`.

## Вариант 2 — вручную по SSH

```bash
ssh root@78.111.91.152
mkdir -p /opt/cppk && cd /opt/cppk
# загрузите файлы app.py data.py xls_import.py page_ui.html (scp/git/файловый менеджер панели)

python3 -m pip install fastapi "uvicorn[standard]" openpyxl xlrd
PORT=9000 nohup python3 /opt/cppk/app.py > /var/log/cppk.log 2>&1 &
```

Для автозагрузки через systemd:

```bash
cp /opt/cppk/cppk.service /etc/systemd/system/
systemctl daemon-reload && systemctl enable --now cppk
```

Если на VPS стоит хостинг-панель (ISPmanager/FastPanel) — можно создать
Python-приложение с точкой входа `app:app` и портом 9000 вместо systemd.

## Проверка после размещения

* `curl http://78.111.91.152:9000/api/health` → `{"status":"ok",...}`
* Открыть с телефона `http://78.111.91.152:9000/` — загрузить XLS через форму,
  убедиться что объекты находятся поиском (город, № билетомата, IP).
* Логи: `journalctl -u cppk -f`

## Частые проблемы

| Симптом | Причина / решение |
|---|---|
| Порт 22 открыт, но страница не грузится | Неверный пароль/ключ SSH — уточните учётные данные у хостинга; либо порт 9000 закрыт в фаерволе **панели управления VPS** (в Тильда-панели Selectel, dgis-панели и т.п. нужно добавить правило) |
| `.xls` не импортируется | На сервере нет `xlrd`: `pip install xlrd` |
| Нужно обновить код | Перезапустить `./deploy.sh` или `systemctl restart cppk` |

**Важно:** сервис работает по HTTP без авторизации — база настроек будет видна
любому, кто знает адрес. Для продакшена рекомендуется поставить nginx с basic-auth
или ограничить доступ по IP.
