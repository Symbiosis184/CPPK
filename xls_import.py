# -*- coding: utf-8 -*-
"""Импорт XLS/XLSX с объектной базой билетоматов в веб-сервис.

Поддерживаются «железнодорожные» таблицы вида:
  № п/п | Направление | Участок | Название станции | Модель БПА | Зав. № | ID БПА |
  Модель pos-терминала | s/n Pin-Pad | TID Pin-Pad СБЕР | Адрес регистрации |
  Место расчетов | Тип связи (СПД/GSM) | ... | IP БПА | IP PinPad | IP ФР | Шлюз | Маска

Столбцы определяются по заголовкам (порядок не важен), распознаются кириллица
и латиница, разные варианты названий колонок. Неизвестные колонки сохраняются
как «прочее» и участвуют в поиске.
"""
import datetime
import re

from openpyxl import load_workbook

try:  # старый .xls (BIFF) — через xlrd, если установлен
    import xlrd  # noqa: F401
    HAS_XLRD = True
except Exception:
    HAS_XLRD = False


# ---------------------------------------------------------------------------
# Словарь заголовков: внутреннее имя -> список вариантов названия колонки
# ---------------------------------------------------------------------------
_ALIASES = {
    "number":      ["id бпа", "id apb", "id билета", "id билетомата",
                    "номер билетомата", "№ билетомата", "номер бма", "номер бпа",
                    "номер apb", "номер tm", "номер автомата", "номер устройства",
                    "инвентарный номер", "инв номер", "id терминала",
                    "ticket machine id"],
    "serial":      ["зав №", "зав номер", "зав. номер", "заводской номер",
                    "serial number", "s/n", "sn"],
    "station":     ["название станции", "станция", "остановочный пункт",
                    "наименование станции", "station"],
    "direction":   ["направление", "напрвление", "direction"],
    "section":     ["участок", "секция", "section"],
    "tm_model":    ["модель бпа", "модель bpa", "модель apb",
                    "модель абп", "модель билетомата", "модель бма", "модель tm",
                    "модель автомата", "тип бпа", "тип билетомата", "tm model"],
    "ticket_machine": ["тип билетомата", "вид билетомата", "система билетомата",
                       "тип tm"],
    "model":       ["модель pos терминала", "pos терминал", "pos терминала",
                    "модель терминала", "модель пктс", "модель пкт",
                    "пктс", "пкт", "банк терминал", "банковский терминал",
                    "terminal model", "терминал"],
    "vendor":      ["вендор", "производитель", "vendor", "brand", "марка"],
    "ip":          ["ip бпа", "ip bpa", "ip apb", "ip абп", "ip билетомата",
                    "ip бма", "ip терминала", "ip адрес", "адрес ip",
                    "ipv4", "ip address", "ip"],
    "pinpad_ip":   ["ip pinpad", "ip pin pad", "ip пинпад"],
    "fr_ip":       ["ip фр", "ip ккм", "ip фискального", "ip фн"],
    "mask":        ["маска", "маска сети", "mask", "subnet", "prefix"],
    "gateway":     ["шлюз", "основной шлюз", "gateway", "default gateway", "gw"],
    "dns":         ["dns", "днс", "dns сервер", "dns servers", "резервный dns"],
    "vlan":        ["vlan", "влан", "vrf"],
    "connection":  ["тип связи", "канал связи", "способ связи", "связь",
                    "тип подключения", "connection type", "link type"],
    "scheme":      ["dhcp", "статический", "static", "ip dhcp", "получение ip",
                    "способ получения ip", "ip схема", "ip scheme", "адресация"],
    "server":      ["сервер", "host", "хост", "серверы", "servers", "эквайринг",
                    "адрес сервера", "url"],
    "ports":       ["порты", "port", "ports", "tcp порты"],
    "interface":   ["com порт", "rs232", "rs 232", "rs485", "rs 485", "usb",
                    "ethernet", "eth", "интерфейс", "соединение", "interface",
                    "baud", "скорость", "последовательный"],
    "sim":         ["сим", "sim карта", "мкс", "мобильные операторы",
                    "оператор", "apn", "imei"],
    "firmware":    ["прошивка", "fw", "firmware", "версия по", "version"],
    "reg_address": ["адрес регистрации для сбер банка", "адрес регистрации",
                    "юридический адрес", "адрес регистрации сбер"],
    "place":       ["место расчетов для сбер банка", "место расчетов",
                    "место установки", "месторасположение", "точка", "объект",
                    "location"],
    "mcd":         ["прин сть к мцд", "принадлежность к мцд", "мцд"],
    "region":      ["прин сть", "принадлежность", "регион", "филиал", "рц"],
    "tid":         ["tid pin pad сбер", "tid pin pad", "tid pinpad", "tid",
                    "terminal id"],
    "sn_pinpad":   ["s n pin pad", "s n pinpad", "sn pinpad"],
    "city":        ["город", "city", "населенный пункт", "нп"],
    "address":     ["адрес объекта", "фактический адрес", "address", "адрес"],
    "notes":       ["примечание", "примечания", "комментарий", "note", "notes",
                    "remarks", "особенности", "доп", "дополнительно"],
    "install_date": ["дата установки", "дата монтажа", "installed",
                     "введен в эксплуатацию"],
}


def _norm_head(s: str) -> str:
    s = str(s).strip().lower()
    s = s.replace("‘", "'").replace("’", "'").replace("`", "'")
    s = s.replace("“", '"').replace("”", '"')
    s = s.replace("\u00a0", " ")
    for d in ("\u2010", "\u2011", "\u2012", "\u2013", "\u2014"):
        s = s.replace(d, "-")
    s = s.replace("№", "no ")
    s = re.sub(r"[^a-zа-я0-9]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def _build_lookup():
    lookup = {}
    for key, variants in _ALIASES.items():
        for v in variants:
            nv = _norm_head(v)
            if nv:
                lookup[nv] = key
    return lookup


_LOOKUP = _build_lookup()


def _score_header(h: str, alias: str) -> float:
    """Насколько заголовок похож на вариант имени колонки."""
    if not h or not alias:
        return 0.0
    if h == alias:
        return 4.0
    if h.startswith(alias + " ") or alias.startswith(h + " "):
        return 3.0
    if alias in h or h in alias:
        return 2.0
    hw, aw = h.split(), alias.split()
    if hw and aw and hw[0] == aw[0] and len(hw[0]) >= 4:
        return 1.0
    return 0.0


def _match_column(h: str):
    best_key, best_sc = None, 0.0
    for alias, key in _LOOKUP.items():
        sc = _score_header(h, alias)
        if sc > best_sc:
            best_key, best_sc = key, sc
    return best_key, best_sc


_HEADER_HINTS = ("no", "модель", "model", "ip", "шлюз", "маска", "станц",
                 "адрес", "тип", "tid", "зав", "id", "прин", "связь",
                 "направ", "участок", "сервер", "dns", "vlan", "порт",
                 "номер", "город", "терминал", "pin")


def _detect_header_row(rows):
    best_row, best_map, best_score = None, {}, 0.0
    for ri, row in enumerate(rows[:10]):
        cells = [c for c in row if c is not None and str(c).strip()]
        if len(cells) < 3:
            continue
        mapping, score, hits = {}, 0.0, 0
        for ci, c in enumerate(row):
            if c is None:
                continue
            h = _norm_head(c)
            if not h or len(h) > 60:
                continue
            key, sc = _match_column(h)
            if key and sc >= 2.0 and key not in mapping.values():
                mapping[ci] = key
                score += sc
                hits += 1
        if hits >= 3 and score > best_score:
            best_row, best_map, best_score = ri, mapping, score
    return best_row, best_map


def _find_header(rows):
    hdr_i, colmap = _detect_header_row(rows)
    if hdr_i is not None:
        return hdr_i, colmap
    for ri, row in enumerate(rows[:6]):
        hints = sum(1 for c in row if c is not None and
                    any(x in _norm_head(str(c)) for x in _HEADER_HINTS))
        if hints >= 3:
            _, colmap = _detect_header_row([row])
            return ri, colmap
    return -1, {}


def _cell_to_str(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.strftime("%d.%m.%Y")
    return str(v).strip()


def _guess_scheme(val: str) -> str:
    v = (val or "").lower()
    if any(x in v for x in ("dhcp", "динам", "автомат")):
        if any(x in v for x in ("стат", "static", "ручн")):
            return "dhcp+static"
        return "dhcp"
    if any(x in v for x in ("static", "стат", "ручн", "фикс")):
        return "static"
    return ""


def _guess_type(val: str) -> str:
    v = (val or "").upper()
    if "BPA" in v or "БПА" in v or "БИЛЕТ" in v:
        return "BPA"
    if "APB" in v or "АБР" in v or "ПАРК" in v or "ПЛАТН" in v:
        return "APB"
    return v.strip()


def _looks_like_ip(v: str) -> bool:
    return bool(re.fullmatch(r"(\d{1,3}\.){3}\d{1,3}", v))


_IP_RE = re.compile(r"\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b")


def _extract_missing(rec: dict, raw_cells):
    joined = " ".join(_cell_to_str(c) for c in raw_cells)
    ips = [i for i in _IP_RE.findall(joined) if _looks_like_ip(i)]
    if not rec.get("ip") and ips:
        rec["ip"] = ips[0]
    if not rec.get("mask"):
        m = re.search(r"(?:маска|mask)[\s:]*(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})",
                      joined, re.I)
        if m:
            rec["mask"] = m.group(1)
    if not rec.get("gateway"):
        g = re.search(r"(?:шлюз|gateway|gw)[\s:]*(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})",
                      joined, re.I)
        if g:
            rec["gateway"] = g.group(1)
    if not rec.get("vlan"):
        vl = re.search(r"vlan[\s:\-]*(\d{1,4})", joined, re.I)
        if vl:
            rec["vlan"] = vl.group(1)
    return rec


def _parse_sheet(name: str, rows):
    rows = [r for r in rows if any(c is not None and str(c).strip() for c in r)]
    if not rows:
        return []
    hdr_i, colmap = _find_header(rows)
    width = max(len(r) for r in rows)
    header_row = rows[hdr_i] if hdr_i >= 0 else []
    records = []
    for row in rows[hdr_i + 1:]:
        cells = [_cell_to_str(c) for c in row] + [""] * (width - len(row))
        rec = {k: "" for k in _ALIASES}
        extras = []
        for ci, val in enumerate(cells):
            key = colmap.get(ci)
            if key:
                if val and not rec[key]:
                    rec[key] = val
            elif val:
                head = _cell_to_str(header_row[ci]) if ci < len(header_row) else ""
                extras.append(f"{head}: {val}" if head else val)
        rec["_extras"] = "; ".join(extras)[:400]
        if hdr_i < 0 or len(colmap) < 3:
            for v in cells:
                if not v:
                    continue
                if _looks_like_ip(v) and not rec["ip"]:
                    rec["ip"] = v
                elif re.search(r"[А-Яа-яA-Za-z]+\s*\d+\-\d+", v) and not rec["tm_model"]:
                    rec["tm_model"] = v
        rec = _extract_missing(rec, row)
        if not (rec.get("ip") or rec.get("number") or rec.get("station")
                or rec.get("model") or rec.get("tm_model") or rec.get("city")):
            continue
        rec["_sheet"] = name
        records.append(rec)
    return records


def parse_xlsx(path: str):
    wb = load_workbook(path, data_only=True, read_only=True)
    all_records, sheets_used = [], []
    for ws in wb.worksheets:
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        recs = _parse_sheet(ws.title, rows)
        if recs:
            sheets_used.append((ws.title, len(recs)))
        all_records.extend(recs)
    wb.close()
    return all_records, sheets_used


def parse_xls(path: str):
    import xlrd
    book = xlrd.open_workbook(path)
    records, sheets_used = [], []
    for sh in book.sheets():
        grid = []
        for r in range(sh.nrows):
            grid.append([sh.cell_value(r, c) if sh.cell_value(r, c) != "" else None
                         for c in range(sh.ncols)])
        recs = _parse_sheet(sh.name, grid)
        if recs:
            sheets_used.append((sh.name, len(recs)))
        records.extend(recs)
    return records, sheets_used


def parse_file(path: str):
    low = path.lower()
    if low.endswith((".xlsx", ".xlsm")):
        return parse_xlsx(path)
    if low.endswith(".xls"):
        if HAS_XLRD:
            return parse_xls(path)
        raise ValueError("Для .xls нужен пакет xlrd (pip install xlrd) или пересохраните в .xlsx")
    raise ValueError("Поддерживаются файлы .xls и .xlsx")


# ---------------------------------------------------------------------------
# Преобразование записи файла -> карточка объекта для фронтенда
# ---------------------------------------------------------------------------
def to_card(rec: dict, idx: int):
    num = (rec.get("number") or rec.get("serial") or rec.get("tid")
           or f"N-{idx:04d}").strip()
    # тип билетомата: явная колонка «тип», иначе по модели БПА (БПА-* -> BPA)
    raw_tm = rec.get("ticket_machine") or rec.get("type") or ""
    ttype = _guess_type(raw_tm) or _guess_type(rec.get("tm_model", ""))
    station = rec.get("station", "")
    city = rec.get("city", "")
    reg_addr = rec.get("reg_address", "")
    place = rec.get("place", "")
    address = rec.get("address", "") or place or reg_addr
    conn = rec.get("connection", "")
    if not city:
        m = re.search(r"\b(?:г\.|город)\s*([А-ЯЁ][а-яё]+)", reg_addr or address)
        if m and m.group(1).lower() not in ("о", "обл"):
            city = m.group(1)
        elif "москв" in (reg_addr + " " + address).lower():
            city = "Москва"
    card = {
        "id": f"OBJ-{num}",
        "source": "upload",
        "object": True,
        "number": num,
        "serial": rec.get("serial", ""),
        "city": city,
        "station": station,
        "direction": rec.get("direction", ""),
        "section": rec.get("section", ""),
        "address": address,
        "reg_address": reg_addr,
        "place": place,
        "region": rec.get("region", ""),
        "mcd": rec.get("mcd", ""),
        "ticket_machine": ttype or "—",
        "tm_model": rec.get("tm_model", ""),
        "model": rec.get("model", "") or "—",
        "vendor": rec.get("vendor", ""),
        "ip": rec.get("ip", ""),
        "pinpad_ip": rec.get("pinpad_ip", ""),
        "fr_ip": rec.get("fr_ip", ""),
        "mask": rec.get("mask", ""),
        "gateway": rec.get("gateway", ""),
        "dns": rec.get("dns", ""),
        "vlan": rec.get("vlan", ""),
        "ip_scheme": _guess_scheme(rec.get("scheme", "")) or ("static" if rec.get("ip") else ""),
        "connection": conn,
        "server": rec.get("server", ""),
        "ports_out": rec.get("ports", ""),
        "com_port": rec.get("interface", ""),
        "sim": rec.get("sim", ""),
        "tid": rec.get("tid", ""),
        "sn_pinpad": rec.get("sn_pinpad", ""),
        "auth": "",
        "firmware": rec.get("firmware", ""),
        "setup_steps": [],
        "notes": rec.get("notes", ""),
        "extras": rec.get("_extras", ""),
        "install_date": rec.get("install_date", ""),
        "tags": [x for x in [city, station, rec.get("direction", ""),
                             rec.get("section", ""), ttype, num,
                             rec.get("_sheet", "")] if x],
    }
    return card
