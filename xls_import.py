# -*- coding: utf-8 -*-
"""Импорт XLS/XLSX с объектной базой билетоматов в веб-сервис.

Файл инженера содержит строки вида: город, адрес, номер билетомата,
тип (BPA/APB), модель терминала и его сетевые настройки.
Модуль определяет столбцы по «говорящим» заголовкам (порядок колонок
не важен, поддерживаются оба языка заголовков) и приводит каждую строку
к единому формату карточки объекта.
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
    "city":        ["город", "city", "населенный пункт", "нп"],
    "address":     ["адрес", "address", "место установки", "точка", "объект",
                    "место", "location"],
    "number":      ["номер билетомата", "№ билетомата", "номер бма", "номер tm",
                    "номер билета", "номер машины", "номер автомата", "id билетомата",
                    "номер устройства", "номер", "№", "id", "number", "bma number",
                    "ticket machine number", "инв", "инвентарный номер"],
    "type":        ["тип", "тип билетомата", "tm", "вид", "type", "система",
                    "тип терминала продажи", "класс"],
    "tm_model":    ["модель билетомата", "модель бма", "модель tm", "модель автомата",
                    "tm model", "модель"],
    "model":       ["модель терминала", "терминал", "модель пктс", "модель пкт",
                    "пктс", "пкт", "банк-терминал", "банковский терминал",
                    "terminal", "terminal model", "устройство"],
    "vendor":      ["вендор", "производитель", "vendor", "brand", "марка"],
    "ip":          ["ip", "ip-адрес", "ip адрес", "адрес ip", "ipv4", "ip address"],
    "mask":        ["маска", "mask", "маска сети", "subnet", "prefix"],
    "gateway":     ["шлюз", "gateway", "основной шлюз", "gw", "default gateway"],
    "dns":         ["dns", "днс", "dns сервер", "dns servers", "резервный dns"],
    "vlan":        ["vlan", "влан", "vrf"],
    "scheme":      ["dhcp", "статический", "static", "ip dhcp", "получение ip",
                    "способ получения ip", "ip схема", "ip scheme", "адресация"],
    "server":      ["сервер", "host", "хост", "серверы", "servers", "эквайринг",
                    "адрес сервера", "url"],
    "ports":       ["порты", "port", "ports", "端口", "tcp порты"],
    "connection":  ["com", "rs232", "rs-232", "rs485", "rs-485", "usb", "ethernet",
                    "eth", "интерфейс", "связь", "соединение", "interface", "baud",
                    "скорость", "последовательный"],
    "sim":         ["сим", "sim-карта", "sim карта", "мкс", "мобильные операторы",
                    "оператор", "apn", "imei"],
    "firmware":    ["прошивка", "fw", "firmware", "версия", "version"],
    "notes":       ["примечание", "примечания", "комментарий", "note", "notes",
                    "remarks", "особенности", "доп", "дополнительно"],
    "install_date":["дата установки", "дата монтажа", "installed", "date",
                    "введен в эксплуатацию"],
}


def _norm_head(s: str) -> str:
    s = str(s).strip().lower()
    s = re.sub(r"\s+", " ", s)
    return s


def _build_lookup():
    lookup = {}
    for key, variants in _ALIASES.items():
        for v in variants:
            lookup[_norm_head(v)] = key
    return lookup


_LOOKUP = _build_lookup()


def _score_header(h: str, alias: str) -> float:
    """Насколько заголовок похож на вариант имени колонки."""
    if h == alias:
        return 3.0
    if h.startswith(alias) or alias.startswith(h):
        return 2.0
    if alias in h or h in alias:
        return 1.5
    return 0.0


def _detect_header_row(rows):
    """Возвращает (номер_строки, {col_idx: key}) или (None, None)."""
    best_row, best_map, best_score = None, None, 0
    for ri, row in enumerate(rows[:10]):
        cells = [c for c in row if c is not None and str(c).strip()]
        if len(cells) < 3:
            continue
        mapping, score = {}, 0
        for ci, c in enumerate(row):
            if c is None:
                continue
            h = _norm_head(c)
            if not h or any(ch.isdigit() for ch in h.replace(".0", "")) and len(h) <= 6:
                continue
            matched_key, matched_score = None, 0
            for alias, key in _LOOKUP.items():
                sc = _score_header(h, alias)
                if sc > matched_score:
                    matched_key, matched_score = key, sc
            if matched_key and matched_score >= 1.5:
                if matched_key not in mapping.values():
                    mapping[ci] = matched_key
                    score += matched_score
        # считаем годным, если распознали хотя бы номер/город/ip
        recognized = set(mapping.values())
        core = recognized & {"number", "city", "ip", "model", "address"}
        if score >= 4 and len(core) >= 2 and (score > best_score):
            best_row, best_map, best_score = ri, mapping, score
    # если в первой строке уже есть IP-адрес или тип BPA/APB — это данные, не заголовок
    if rows:
        first = " ".join(_cell_to_str(c) for c in rows[0])
        if _IP_RE.search(first) or re.search(r"\b(BPA|APB)\b", first):
            return -1, {}
    return best_row, best_map


def _cell_to_str(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.strftime("%d.%m-%Y" if hasattr(v, "strftime") else "%Y-%m-%d")
    return str(v).strip()


def _guess_scheme(val: str) -> str:
    v = val.lower()
    if any(x in v for x in ("dhcp", "динам", "автомат")):
        if any(x in v for x in ("стат", "static", "ручн")):
            return "dhcp+static"
        return "dhcp"
    if any(x in v for x in ("static", "стат", "ручн", "фикс")):
        return "static"
    return ""


def _guess_type(val: str) -> str:
    v = val.upper()
    if "BPA" in v or "БИЛЕТ" in v or "GSP" in v or "ТРАНСП" in v:
        return "BPA"
    if "APB" in v or "ПАРК" in v or "ПАРКОМ" in v or "Платн" in v:
        return "APB"
    return val.strip().upper()


def _looks_like_ip(v: str) -> bool:
    return bool(re.fullmatch(r"(\d{1,3}\.){3}\d{1,3}", v))


_IP_RE = re.compile(r"\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}(?:/\d{1,2})?\b")


def _extract_missing(rec: dict, raw_cells):
    """Если колонки не распознаны — пытаемся вытащить IP/маску/шлюз из ячеек."""
    joined = " ".join(_cell_to_str(c) for c in raw_cells)
    ips = _IP_RE.findall(joined)
    if not rec.get("ip"):
        plain = [i for i in ips if _looks_like_ip(i)]
        if plain:
            rec["ip"] = plain[0]
    if not rec.get("mask"):
        m = re.search(r"mask|маска[:\s]*((\d{1,3}\.){3}\d{1,3}|/\d{1,2})", joined, re.I)
        if m:
            rec["mask"] = m.group(1)
        elif "/" in (rec.get("ip") or ""):
            rec["ip"], rec["mask"] = rec["ip"].split("/", 1)
    if not rec.get("gateway"):
        g = re.search(r"(?:gw|шлюз|gateway)[:\s]*((\d{1,3}\.){3}\d{1,3})", joined, re.I)
        if g:
            rec["gateway"] = g.group(1)
    if not rec.get("vlan"):
        vl = re.search(r"vlan[ :\-]*(\d{1,4})", joined, re.I)
        if vl:
            rec["vlan"] = vl.group(1)
    return rec


def parse_xlsx(path: str):
    wb = load_workbook(path, data_only=True, read_only=True)
    all_records, sheets_used = [], []
    for ws in wb.worksheets:
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        rows = [r for r in rows if any(c is not None and str(c).strip() for c in r)]
        if not rows:
            continue
        hdr_i, colmap = _detect_header_row(rows)
        if hdr_i is None:
            # нет заголовков — пробуем первую строку как данные, колонки угадываем
            colmap, hdr_i = {}, -1
        elif not colmap:
            # распознали только служебные колонки (город/адрес) — остальные угадываем по значениям
            pass
        records = []
        for row in rows[hdr_i + 1:]:
            cells = list(row) + [None] * (max(colmap) + 1 if colmap else len(row))
            rec = {k: "" for k in _ALIASES}
            for ci, key in colmap.items():
                if ci < len(row):
                    rec[key] = _cell_to_str(row[ci])
            if not any(rec.values()) and colmap:
                continue
            # недостающие поля угадываем по содержимому ячеек (для файлов без заголовков
            # или с частично распознанными заголовками)
            if len(colmap) < 4:
                for c in cells:
                    v = _cell_to_str(c)
                    if not v:
                        continue
                    if _looks_like_ip(v) and not rec["ip"]:
                        rec["ip"] = v
                    elif re.fullmatch(r"БМА?[-–\s]?\d{3,}|№?\s?\d{5,}", v, re.I) and not rec["number"]:
                        rec["number"] = v
                    elif v.upper() in ("BPA", "APB") and not rec["type"]:
                        rec["type"] = v
                    elif re.search(r"\d{1,3}\.\d{1,3}\.\d{1,3}", v) or re.search(r"(RS|COM|USB|ETH)", v, re.I):
                        if not rec["connection"] and not _looks_like_ip(v):
                            rec["connection"] = v
                    elif re.search(r"[A-Za-zА-Яа-я]{3,}.*\d|\d.*[A-Za-zА-Яа-я]", v) and \
                            not rec["model"] and not rec["address"] and not rec["city"] \
                            and len(v) < 40 and "," not in v:
                        rec["model"] = v
            rec = _extract_missing(rec, cells)
            if not (rec.get("ip") or rec.get("number") or rec.get("city")
                    or rec.get("model")):
                continue
            rec["_sheet"] = ws.title
            records.append(rec)
        if records:
            sheets_used.append((ws.title, len(records)))
        all_records.extend(records)
    wb.close()
    return all_records, sheets_used


def parse_xls(path: str):
    import xlrd
    book = xlrd.open_workbook(path)
    rows_all, sheets_used, records = [], [], []
    for sh in book.sheets():
        grid = []
        for r in range(sh.nrows):
            grid.append([sh.cell_value(r, c) if sh.cell_value(r, c) != "" else None
                         for c in range(sh.ncols)])
        rows = [r for r in grid if any(c is not None and str(c).strip() for c in r)]
        if not rows:
            continue
        hdr_i, colmap = _detect_header_row(rows)
        if hdr_i is None:
            colmap, hdr_i = {}, -1
        cnt = 0
        for row in rows[hdr_i + 1:]:
            rec = {k: "" for k in _ALIASES}
            for ci, key in colmap.items():
                if ci < len(row):
                    rec[key] = _cell_to_str(row[ci])
            if not any(rec.values()):
                continue
            rec = _extract_missing(rec, row)
            if not (rec.get("ip") or rec.get("number") or rec.get("city")):
                continue
            rec["_sheet"] = sh.name
            records.append(rec)
            cnt += 1
        sheets_used.append((sh.name, cnt))
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
_FIELD_LABELS = [
    ("ip", "IP-адрес"), ("mask", "Маска"), ("gateway", "Шлюз"),
    ("dns", "DNS"), ("vlan", "VLAN"), ("scheme", "Схема IP"),
    ("server", "Серверы"), ("ports", "Порты"), ("connection", "Связь с БМА"),
    ("sim", "SIM/APN"), ("firmware", "Прошивка"),
    ("install_date", "Дата установки"), ("notes", "Примечания"),
]


def to_card(rec: dict, idx: int):
    num = rec.get("number") or f"N-{idx:04d}"
    ttype = _guess_type(rec.get("type", ""))
    model = rec.get("model") or rec.get("tm_model") or "Терминал (см. примечания)"
    vendor = rec.get("vendor") or ""
    scheme = _guess_scheme(rec.get("scheme", "")) or ("static" if rec.get("ip") else "")
    card = {
        "id": f"OBJ-{num}",
        "source": "upload",
        "object": True,
        "number": num,
        "city": rec.get("city", ""),
        "address": rec.get("address", ""),
        "ticket_machine": ttype or "—",
        "tm_model": rec.get("tm_model", ""),
        "model": model,
        "vendor": vendor,
        "ip": rec.get("ip", ""),
        "mask": rec.get("mask", ""),
        "gateway": rec.get("gateway", ""),
        "dns": rec.get("dns", ""),
        "vlan": rec.get("vlan", ""),
        "ip_scheme": scheme,
        "server": rec.get("server", ""),
        "ports_out": rec.get("ports", ""),
        "connection": rec.get("connection", ""),
        "com_port": rec.get("connection", ""),
        "sim": rec.get("sim", ""),
        "auth": "",
        "firmware": rec.get("firmware", ""),
        "setup_steps": [],
        "notes": rec.get("notes", ""),
        "install_date": rec.get("install_date", ""),
        "tags": [x for x in [rec.get("city", ""), ttype, num, rec.get("_sheet", "")] if x],
    }
    return card
