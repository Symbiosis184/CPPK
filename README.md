# CPPK — сетевые настройки билетоматов BPA/APB (СберСервис)

Веб-сервис для разъездных инженеров: поиск по сетевым настройкам банковских
терминалов под билетоматы + загрузка XLS/XLSX с адресной базой (город, №
билетомата, IP/VLAN/шлюзы). Мобильный интерфейс, живой поиск.

## Быстрая установка на VPS (с GitHub)

1. Создайте **публичный** репозиторий на GitHub и загрузите в него файлы проекта:

```bash
cd /путь/к/проекту
git init -b main
git add app.py data.py xls_import.py page_ui.html requirements.txt cppk.service \
        setup_from_git.sh install.sh .gitignore DEPLOY.md README.md
git commit -m "CPPK initial"
git remote add origin https://github.com/ВАШ_ЛОГИН/cppk.git
git push -u origin main
```

2. На VPS (root@78.111.91.152) — **одна команда**:

```bash
curl -fsSL https://raw.githubusercontent.com/ВАШ_ЛОГИН/cppk/main/setup_from_git.sh | bash
```

(скрипт сам определит URL репозитория, из которого его скачали; можно задать явно:
`REPO_URL=https://github.com/ВАШ_ЛОГИН/cppk.git bash setup_from_git.sh`)

Что делает скрипт:
- клонирует репозиторий в `/opt/cppk` (или обновляет `git reset --hard` при повторном запуске);
- ставит зависимости в venv (обход блокировки PEP 668 в Debian 13);
- регистрирует systemd-службу `cppk` (автозапуск + перезапуск при падении), порт 9000;
- открывает порт в ufw, проверяет `http://127.0.0.1:9000/api/health`.

Готово: **http://78.111.91.152:9000**

## Обновление после изменений на GitHub

Повторно выполните ту же команду установки — `setup_from_git.sh` подтянет новый
код из git, обновит venv и перезапустит службу.

## Управление

```bash
systemctl status|restart|stop cppk
journalctl -u cppk -f          # логи
```

## API
- `GET /` — веб-интерфейс (мобильная версия)
- `POST /api/upload` — загрузка XLS/XLSX
- `GET /api/objects?q=&city=` — поиск по объектам
- `GET /api/terminals?q=&tm=&vendor=` — справочник терминалов
- `GET /api/health`, Swagger: `/docs`

## Если страница не открывается снаружи
Откройте порт 9000 в фаерволе панели хостинга VPS (Selectel/Timeweb/dgis) —
локальный `curl http://127.0.0.1:9000/api/health` при этом может быть OK.
