# -*- coding: utf-8 -*-
"""Импорт XLS/XLSX с объектной базой билетоматов (БПА/APB) в веб-сервис.

Формат файла инженера (ж/д билеты, Сбер):
  № п/п | Направление | Участок | Название станции | Модель БПА | Зав. № | ID БПА |
  Модель pos-терминала | s/n Pin-Pad | TID Pin-Pad СБЕР | Адрес регистрации |
  Место расчетов | Тип связи (СПД/GSM) | Прин-сть к МЦД | Прин-сть к Москве/МО |
  IP БПА | IP PinPad | IP ФР | Шлюз | Маска

Столбцы определяются по «говорящим» заголовкам (порядок не важен).
Значения берутся СТРОГО из соответствующих колонок — никаких догадок,
чтобы данные оставались точными как в файле.
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
# Внутренние ключи -> варианты заголовков колонки
# ---------------------------------------------------------------------------
_ALIASES = {
    "pasp":        ["№ п/п", "n п/п", "н п/п", "пп", "№п/п", "п/п", "no", "номер строки"],
    "direction":   ["направление", "напрвление", "направлени", "direction"],
    "section":     ["участок", "yчасток", "сектор", "section"],
    "station":     ["название станции", "станция", "остановочный пункт",
                    "ноп", "station", "название остановочного пункта"],
    "tm_model":    ["модель бпа", "модель bppa", "модель билетомата", "модель бма",
                    "модель апб", "модель tm", "tm model", "модель автомата"],
    "zav":         ["зав. №", "зав №", "зав.", "заводской номер", "заводской н",
                    "заводской", "ser", "serial number"],
    "number":      ["id бпа", "ид бпа", "номер билетомата", "№ билетомата",
                    "номер бма", "id бма", "номер tm", "id tm", "номер автомата",
                    "ticket machine number", "bma number"],
    "model":       ["модель pos-терминала", "модель pos терминала", "pos-терминал",
                    "pos терминал", "модель терминала", "терминал", "модель пктс",
                    "модель пкт", "пктс", "пкт", "terminal", "terminal model",
                    "банк-терминал", "банковский терминал", "устройство"],
    "sn_pinpad":   ["s/n pin-pad", "sn pin pad", "s/n pinpad", "серийный номер pin-pad",
                    "серийник pin-pad", "pin-pad sn", "sn pin_pad", "s/n пинпад"],
    "tid":         ["tid pin-pad сбер", "tid pin-pad", "tid pinpad", "tid",
                    "tid пинпад"],
    "city":        ["город", "населенный пункт", "город", "city", "муниципальный округ"],
    "reg_address": ["адрес регистрации для сбер банка", "адрес регистрации",
                    "адрес регистр", "address"],
    "address":     ["адрес", "место установки адрес", "адрес объекта"],
    "place":       ["место расчетов для сбер банка", "место расчетов", "место расчётов",
                    "место установки", "место", "location"],
    "connection":  ["тип связи", "тип связи (спд / gsm)", "тип связи (спд/gsm)",
                    "канал связи", "способ связи", "связь"],
    "mcd":         ["прин-сть к мцд", "принадлежность к мцд", "мцд"],
    "region":      ["прин-сть к москве/мо/рег. центры", "прин-сть к москве/мо",
                    "прин-сть к москве", "прин-сть к мо", "принадлежность",
                    "прин-сть", "регион", "зона"],
    "ip":          ["ip бпа", "ип бпа", "ip билетомата", "ip бма", "ip адрес бпа",
                    "ip-адрес", "ip адрес", "ip", "ipv4", "ip address"],
    "ip_pinpad":   ["ip pinpad", "ip pin-pad", "ип pin-pad", "ip пинпада",
                    "ip pin_pad", "ip pos-терминала", "ip pos терминала"],
    "ip_fr":       ["ip фр", "фр ip", "ip фискальный регистратор", "ip ккт"],
    "gateway":     ["шлюз", "основной шлюз", "gateway", "gw", "default gateway"],
    "mask":        ["маска", "маска сети", "mask", "subnet", "prefix"],
    "dns":         ["dns", "днс", "dns сервер", "dns servers", "резервный dns"],
    "vlan":        ["vlan", "влан", "vrf"],
    "scheme":      ["dhcp", "статический", "static", "ip dhcp", "получение ip",
                    "способ получения ip", "ip схема", "ip scheme", "адресация"],
    "server":      ["сервер", "host", "хост", "серверы", "servers", "эквайринг",
                    "адрес сервера", "url"],
    "ports":       ["порты", "port", "ports", "tcp порты"],
    "sim":         ["сим", "sim-карта", "sim карта", "мкс", "мобильные операторы",
                    "оператор", "apn", "imei"],
    "firmware":    ["прошивка", "fw", "firmware", "версия прошивки", "version"],
    "notes":       ["примечание", "примечания", "комментарий", "note", "notes",
                    "remarks", "особенности", "доп", "дополнительно"],
    "install_date": ["дата установки", "дата монтажа", "installed", "date",
                     "введен в эксплуатацию"],
    "vendor":      ["вендор", "производитель", "vendor", "brand", "марка"],
}


def _strict_head(s: str) -> str:
    """Нормализация заголовка: регистр, пробелы, дефисы, частые опечатки."""
    s = str(s).strip().lower()
    s = s.replace("ё", "е")
    s = re.sub(r"[‐‑–—]", "-", s)          # разные тире
    s = re.sub(r"\s+", " ", s)
    s = s.replace("№", "n").replace("#", "n")
    s = re.sub(r"\s*([/.])\s*", r"\1", s)  # "зав. №" -> "зав.n"
    s = re.sub(r"[^a-zа-я0-9/. -]+", "", s)
    s = re.sub(r"\s+", " ", s).strip()
    for a, b in (("bppa", "бпа"), ("pin—pad", "pin-pad")):
        s = s.replace(a, b)
    return s


def _build_exact():
    exact = {}
    for key, variants in _ALIASES.items():
        for v in variants:
            hv = _strict_head(v)
            if hv not in exact:
                exact[hv] = key
    return exact


_EXACT = _build_exact()


def _score_header(h: str, alias: str) -> float:
    if h == alias:
        return 3.0
    if len(alias) >= 4 and (h.startswith(alias) or alias.startswith(h)):
        return 2.0
    if len(alias) >= 5 and (alias in h or h in alias):
        return 1.5
    return 0.0


def _map_header(row):
    """{col_idx: key}; точные совпадения имеют приоритет над нечёткими."""
    mapping, used_keys = {}, set()
    for ci, c in enumerate(row):
        if c is None:
            continue
        key = _EXACT.get(_strict_head(c))
        if key and key not in used_keys:
            mapping[ci] = key
            used_keys.add(key)
    for ci, c in enumerate(row):
        if ci in mapping or c is None:
            continue
        h = _strict_head(c)
        if not h or len(h) > 60:
            continue
        best_key, best_sc = None, 0.0
        for key, variants in _ALIASES.items():
            if key in used_keys:
                continue
            for v in variants:
                sc = _score_header(h, _strict_head(v))
                if sc > best_sc:
                    best_key, best_sc = key, sc
        if best_key and best_sc >= 1.5:
            mapping[ci] = best_key
            used_keys.add(best_key)
    return mapping


def _detect_header_row(rows):
    best_row, best_map, best_score = None, None, -1
    for ri, row in enumerate(rows[:10]):
        cells = [c for c in row if c is not None and str(c).strip()]
        if len(cells) < 3:
            continue
        mapping = _map_header(row)
        score = len(mapping)
        core = set(mapping.values()) & {"number", "station", "ip", "model",
                                        "address", "city", "tid"}
        if score >= 3 and len(core) >= 1 and score > best_score:
            best_row, best_map, best_score = ri, mapping, score
    return best_row, best_map


def _cell_to_str(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    if isinstance(v, (datetime.datetime, datetime.date)):
        try:
            return v.strftime("%d.%m.%Y")
        except Exception:
            return str(v)
    return str(v).strip()


def _looks_like_ip(v: str) -> bool:
    return bool(re.fullmatch(r"(\d{1,3}\.){3}\d{1,3}", v))


_CITY_WORDS = {"г", "город", "г-к", "п", "пгт", "пос", "поселок", "деревня",
               "село", "ст", "хутор", "рп", "кп", "дп", "посёлок", "дер", "сел"}

# Округи/районы МО -> город для поиска
_MO_MAP = {
    "одинцовский": "Одинцово", "рузский": "Руза",
}


def _extract_city(address: str, station: str = "") -> str:
    """Город из адреса регистрации ('143081, Московская обл., г.о. Одинцово , ...')."""
    addr = (address or "").strip()
    if addr:
        found = re.findall(
            r"(?:^|[,;(]\s*)(?:г\.?\s*о\.?|город|г\.|рп|пгт)\s*"
            r"([A-Za-zА-Яа-я][\w\- ]{2,})", addr)
        cands = []
        for f in found:
            c = f.strip().rstrip(",.;: ").lower()
            c = re.split(r"\s+(?:ул|кв-?л|кв|пл|ш|д|стр|б-р|пер|пр-?кт|мкр|р-?н)\b", c)[0].strip()
            if c and c not in _CITY_WORDS:
                cands.append(c)
        if re.search(r"(^|[,(]\s*)г\.?\s*москва\b", addr, re.I):
            return "Москва"
        for c in cands:
            if c in _MO_MAP:
                return _MO_MAP[c]
        if cands:
            return cands[0].capitalize()
    st = (station or "").strip()
    if st:
        return "Москва" if st.lower() == "москва" else st.capitalize()
    return ""


def _guess_scheme(val: str, rec: dict) -> str:
    v = (val or "").lower()
    conn = (rec.get("connection") or "").upper()
    if any(x in v for x in ("dhcp", "динам", "автомат")):
        return "dhcp+static" if any(x in v for x in ("стат", "static", "ручн")) else "dhcp"
    if any(x in v for x in ("static", "стат", "ручн", "фикс")):
        return "static"
    if conn == "GSM":
        return "GSM"
    if rec.get("ip"):
        return "static"
    return ""


def _guess_type(tm_model: str, fallback: str = "") -> str:
    v = (tm_model or "").upper()
    if "БПА" in v or "BPA" in v:
        return "BPA"
    if "АПБ" in v or "APB" in v:
        return "APB"
    fv = (fallback or "").upper()
    if "BPA" in fv or "БИЛЕТ" in fv:
        return "BPA"
    if "APB" in fv or "ПАРК" in fv:
        return "APB"
    return (fallback or "").strip().upper() or "BPA"


def _rows_from_xlsx(path):
    wb = load_workbook(path, data_only=True, read_only=True)
    out = []
    for ws in wb.worksheets:
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        rows = [r for r in rows if any(c is not None and str(c).strip() for c in r)]
        if rows:
            out.append((ws.title, rows))
    wb.close()
    return out


def _rows_from_xls(path):
    import xlrd
    book = xlrd.open_workbook(path)
    out = []
    for sh in book.sheets():
        grid = [[sh.cell_value(r, c) if sh.cell_value(r, c) != "" else None
                 for c in range(sh.ncols)] for r in range(sh.nrows)]
        rows = [r for r in grid if any(c is not None and str(c).strip() for c in r)]
        if rows:
            out.append((sh.name, rows))
    return out


def parse_file(path: str):
    low = path.lower()
    if low.endswith((".xlsx", ".xlsm")):
        sheets = _rows_from_xlsx(path)
    elif low.endswith(".xls"):
        if not HAS_XLRD:
            raise ValueError("Для .xls нужен пакет xlrd (pip install xlrd) или пересохраните в .xlsx")
        sheets = _rows_from_xls(path)
    else:
        raise ValueError("Поддерживаются файлы .xls и .xlsx")

    all_records, sheets_used = [], []
    for sheet_name, rows in sheets:
        hdr_i, colmap = _detect_header_row(rows)
        if hdr_i is None:
            hdr_i, colmap = -1, {}
        records = []
        for row in rows[hdr_i + 1:]:
            rec = {k: "" for k in _ALIASES}
            for ci, key in colmap.items():
                if ci < len(row):
                    rec[key] = _cell_to_str(row[ci])
            meaningful = [v for k, v in rec.items() if v and k != "pasp"]
            if not meaningful:
                continue
            # если заголовков нет вообще — минимальная эвристика по содержимому
            if not colmap:
                ips = [_cell_to_str(c) for c in row if _looks_like_ip(_cell_to_str(c))]
                if not ips:
                    continue
                rec["ip"] = ips[0]
                if len(ips) > 1:
                    rec["ip_pinpad"] = ips[1]
                if len(ips) > 2:
                    rec["ip_fr"] = ips[2]
                nums = [_cell_to_str(c) for c in row
                        if re.fullmatch(r"\d{4,6}", _cell_to_str(c))]
                if nums:
                    rec["number"] = nums[-1]
            rec["_sheet"] = sheet_name
            records.append(rec)
        if records:
            sheets_used.append((sheet_name, len(records)))
        all_records.extend(records)
    return all_records, sheets_used


# ---------------------------------------------------------------------------
# Запись файла -> карточка объекта
# ---------------------------------------------------------------------------
_EXTRA_LABELS = [
    ("direction", "Направление"), ("section", "Участок"), ("zav", "Зав. №"),
    ("sn_pinpad", "s/n Pin-Pad"), ("place", "Место расчетов"),
    ("mcd", "Прин-сть к МЦД"), ("region", "Регион"),
]


def to_card(rec: dict, idx: int):
    num = rec.get("number") or rec.get("zav") or f"N-{idx:04d}"
    ttype = _guess_type(rec.get("tm_model", ""), rec.get("type", ""))
    reg_addr = rec.get("reg_address", "")
    address = reg_addr or rec.get("address", "")
    city = rec.get("city") or _extract_city(address, rec.get("station", ""))
    station = rec.get("station", "")
    conn = rec.get("connection", "")
    scheme = _guess_scheme(rec.get("scheme", ""), rec)
    parts = []
    if station:
        parts.append(f"ст. {station}")
    for key, label in _EXTRA_LABELS:
        if rec.get(key):
            parts.append(f"{label}: {rec[key]}")
    if conn:
        parts.append(f"Тип связи: {conn}")
    if rec.get("ip_pinpad"):
        parts.append(f"IP Pin-Pad: {rec['ip_pinpad']}")
    if rec.get("ip_fr"):
        parts.append(f"IP ФР: {rec['ip_fr']}")
    notes = rec.get("notes", "")
    if parts:
        notes = (notes + "\n" if notes else "") + " · ".join(parts)
    card = {
        "id": f"OBJ-{num}",
        "source": "upload",
        "object": True,
        "number": num,
        "city": city,
        "station": station,
        "address": address,
        "reg_address": reg_addr,
        "direction": rec.get("direction", ""),
        "section": rec.get("section", ""),
        "zav": rec.get("zav", ""),
        "tid": rec.get("tid", ""),
        "sn_pinpad": rec.get("sn_pinpad", ""),
        "ip_pinpad": rec.get("ip_pinpad", ""),
        "ip_fr": rec.get("ip_fr", ""),
        "region": rec.get("region", ""),
        "mcd": rec.get("mcd", ""),
        "ticket_machine": ttype,
        "tm_model": rec.get("tm_model", ""),
        "model": rec.get("model", ""),
        "vendor": rec.get("vendor", ""),
        "ip": rec.get("ip", ""),
        "mask": rec.get("mask", ""),
        "gateway": rec.get("gateway", ""),
        "dns": rec.get("dns", ""),
        "vlan": rec.get("vlan", ""),
        "ip_scheme": scheme,
        "server": rec.get("server", ""),
        "ports_out": rec.get("ports", ""),
        "connection": conn,
        "com_port": "",
        "sim": rec.get("sim", ""),
        "auth": "",
        "firmware": rec.get("firmware", ""),
        "setup_steps": [],
        "notes": notes,
        "install_date": rec.get("install_date", ""),
        "tags": [x for x in [city, station, ttype, num, rec.get("direction", ""),
                             rec.get("section", ""), rec.get("_sheet", "")] if x],
    }
    return card
