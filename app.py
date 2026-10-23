# -*- coding: utf-8 -*-
"""Веб-сервис «Сетевые настройки терминалов BPA/APB» для разъездных инженеров.

Возможности:
  * поиск по встроенному справочнику терминалов и типовым ошибкам;
  * загрузка XLS/XLSX с объектной базой (город, адрес, № билетомата, настройки) —
    карточки объектов сразу доступны для поиска с телефона;
  * сохранение загруженной базы на сервере (upload_store.json) — переживает рестарт;
  * REST API + Swagger на /docs.

Запуск:
    uvicorn app:app --host 0.0.0.0 --port 8000
или:
    python app.py
"""
import json
import os
import time
import re
import unicodedata
from typing import List

from fastapi import Depends, FastAPI, File, Header, Query, UploadFile
from fastapi.responses import HTMLResponse, JSONResponse

from data import TERMINALS, ERRORS
import xls_import

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_STORE = os.path.join(BASE_DIR, "upload_store.json")
META_FILE = os.path.join(BASE_DIR, "upload_meta.json")


def _save_meta(meta: dict):
    with open(META_FILE, "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False, indent=1)


def _load_meta() -> dict:
    try:
        with open(META_FILE, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}

# Пароль на удаление базы (меняется здесь же)
ADMIN_PASSWORD = os.environ.get("CPPK_ADMIN_PASSWORD", "25122023")


def require_admin(x_admin: str = Header(default="")):
    """Доступ к удалению базы — только с правильным паролем (25122023)."""
    from fastapi.exceptions import HTTPException
    if x_admin != ADMIN_PASSWORD:
        raise HTTPException(status_code=403, detail="Неверный пароль")
    return True

app = FastAPI(
    title="CPPK — база сетевых настроек терминалов BPA/APB",
    description="Поиск сетевых настроек банковских терминалов под билетоматы BPA и APB "
                "для разъездных инженеров. Загрузка XLS/XLSX с объектами.",
    version="2.0.0",
)


# ------------------------------------------------------- хранилище объектов ---
def _load_uploaded() -> List[dict]:
    try:
        with open(UPLOAD_STORE, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []


def _save_uploaded(cards: List[dict]):
    with open(UPLOAD_STORE, "w", encoding="utf-8") as f:
        json.dump(cards, f, ensure_ascii=False, indent=1)


UPLOADED = _load_uploaded()  # карточки объектов из загруженного файла


# --------------------------------------------------------------- поиск ------
_CYR2LAT = {
    "а": "a", "б": "b", "в": "v", "г": "g", "д": "d", "е": "e", "ё": "e",
    "ж": "zh", "з": "z", "и": "i", "й": "y", "к": "k", "л": "l", "м": "m",
    "н": "n", "о": "o", "п": "p", "р": "r", "с": "s", "т": "t", "у": "u",
    "ф": "f", "х": "h", "ц": "ts", "ч": "ch", "ш": "sh", "щ": "sch",
    "ъ": "", "ы": "y", "ь": "", "э": "e", "ю": "yu", "я": "ya",
}
_LAT2CYR = {
    "a": "а", "b": "б", "v": "в", "g": "г", "d": "д", "e": "е", "zh": "ж",
    "z": "з", "i": "и", "y": "ы", "k": "к", "l": "л", "m": "м", "n": "н",
    "o": "о", "p": "п", "r": "р", "s": "с", "t": "т", "u": "у", "f": "ф",
    "kh": "х", "h": "х", "ts": "ц", "ch": "ч", "sh": "ш", "sch": "щ",
    "yu": "ю", "ya": "я",
}


def _cyr_to_latin(s: str) -> str:
    return "".join(_CYR2LAT.get(ch, ch) for ch in s)


def _latin_to_cyr(s: str) -> str:
    out, i = [], 0
    while i < len(s):
        for ln in (3, 2, 1):
            chunk = s[i:i + ln]
            if chunk in _LAT2CYR:
                out.append(_LAT2CYR[chunk])
                i += ln
                break
        else:
            out.append(s[i])
            i += 1
    return "".join(out)


_QWERTY2CYR = {
    "q": "й", "w": "ц", "e": "у", "r": "к", "t": "е", "y": "н", "u": "г",
    "i": "ш", "o": "щ", "p": "з", "[": "х", "]": "ъ", chr(92): "ё", "a": "ф",
    "s": "ы", "d": "в", "f": "а", "g": "п", "h": "р", "j": "о", "k": "л",
    "l": "д", ";": "ж", "'": "э", "z": "я", "x": "ч", "c": "с", "v": "м",
    "b": "и", "n": "т", "m": "ь", ",": "б", ".": "ю",
}


def _qwerty_to_cyr(s: str) -> str:
    return "".join(_QWERTY2CYR.get(ch, ch) for ch in s)


def _norm(s: str) -> str:
    s = unicodedata.normalize("NFKC", s.lower().strip())
    return _cyr_to_latin(s)


# Латинские буквы, визуально совпадающие с кириллическими (лат. c -> кир. с и т.п.)
_LAT_LOOKALIKE = {
    "a": "а", "b": "в", "c": "с", "e": "е", "h": "н", "j": "й", "k": "к",
    "m": "м", "o": "о", "p": "р", "s": "с", "t": "т", "x": "х", "y": "у",
}


def _cyr_fix(s: str) -> str:
    """Заменяет латинские буквы-двойники на кирилльские (смешанная раскладка)."""
    return "".join(_LAT_LOOKALIKE.get(ch, ch) for ch in s)


def _variants(q: str):
    q = q.strip().lower()
    v = {q, _cyr_to_latin(q), _cyr_to_latin(_qwerty_to_cyr(q))}
    fixed = _cyr_fix(q)
    if fixed != q:  # в запросе затесались латинские буквы-двойники кириллицы
        v.add(fixed)
        v.add(_cyr_to_latin(fixed))
    if q.isascii():
        cyr = _latin_to_cyr(q)
        v.add(_cyr_to_latin(cyr))
        v.add(cyr)
        # латиница могла быть набранной «не переключив раскладку» кириллицей
        v.add(_cyr_to_latin(_qwerty_to_cyr(cyr)))
    else:
        v |= _variants(_cyr_to_latin(q))
    return {x for x in v if x}


def _search(text: str, needle: str) -> bool:
    tn = _norm(text)
    tl = text.lower()
    return any(v in tn or v in tl for v in _variants(needle))


# Топонимы: город -> дополнительные формы/слова для поиска (в т.ч. район, область)
_PLACE_EXTRA = {
    "москва": "московский москвы московская",
    "пушкино": "пушкинский пушкино пушкинские пушкинская пушкина",
    "одинцово": "одинцовский одинцово одинцовских",
    "руза": "рузский рузский ружский рузы",
    "звенигород": "звенигородский звенигорода",
    "можайск": "можайский можайска",
    "красногорск": "красногорский",
    "подольск": "подольский",
    "сергиев посад": "сергиева посада сергиево-посадский",
}


def _word_forms(w: str) -> set:
    """Варианты слова с учётом транслита/раскладки; для топонима добавляет
    его окончание (Пушкин -> Пушкино) и известные формы (Пушкино -> Пушкинский)."""
    v = _variants(w)
    if len(w) >= 4 and not w.isascii():
        base = w[:-1] if w.endswith(("о", "а", "я", "е")) else w
        for suf in ("о", "ово", "ево", "ино", "ыно", "а", "я"):
            cand = base + suf
            if cand != w:
                v |= _variants(cand)
    extra = _PLACE_EXTRA.get(w)
    if extra:
        for e in extra.split():
            v |= _variants(e)
    return v


def _haystack(t: dict) -> str:
    parts = []
    for key in ("model", "vendor", "ticket_machine", "tm_model", "connection",
                "ip_scheme", "ip", "mask", "gateway", "dns", "server", "com_port",
                "auth", "firmware", "notes", "number", "city", "address",
                "sim", "ports_out", "install_date", "station", "direction",
                "section", "zav", "tid", "sn_pinpad", "ip_pinpad", "ip_fr",
                "region", "mcd", "place", "reg_address"):
        if t.get(key):
            parts.append(str(t[key]))
    if t.get("vlan"):
        parts.append("VLAN " + str(t["vlan"]))
    city = (t.get("city") or "").lower()
    if city in _PLACE_EXTRA:
        parts.append(_PLACE_EXTRA[city])
    parts += list(t.get("tags") or [])
    parts += list(t.get("setup_steps") or [])
    return " ".join(parts)


def filter_all(q: str = "", tm: str = "", vendor: str = "", scheme: str = "",
               city: str = "", source: str = ""):
    """source: '' = всё, 'ref' = справочник, 'obj' = объекты из файла."""
    pool = []
    if source in ("", "ref"):
        pool += TERMINALS
    if source in ("", "obj"):
        pool += UPLOADED
    results = []
    for t in pool:
        if tm and (t.get("ticket_machine") or "").upper() != tm.upper():
            continue
        if vendor and not _search(t.get("vendor", ""), vendor):
            continue
        if scheme and scheme.lower() not in (t.get("ip_scheme") or "").lower():
            continue
        if city and not _search(t.get("city", ""), city):
            continue
        if q:
            hay = _haystack(t)
            hn, hl = _norm(hay), hay.lower()
            if not all(any(v in hn or v in hl for v in _word_forms(w))
                       for w in q.split()):
                continue
        results.append(t)
    # релевантность: точное совпадение номера/города — наверх; объекты из файла первыми
    if q:
        words = [w.lower() for w in q.split()]

        def score(t):
            s = 0
            num = str(t.get("number") or "")
            city = (t.get("city") or "").lower()
            st = (t.get("station") or "").lower()
            for w in words:
                if w == num.lower():
                    s -= 10
                elif city and (w == city or _word_forms(w) & _word_forms(city)):
                    s -= 5
                elif st and st.startswith(w[:4]):
                    s -= 3
            if t.get("object"):
                s -= 1
            return s
        results.sort(key=score)
    return results


# ------------------------------------------------------------------ API ----
@app.get("/api/terminals", response_class=JSONResponse)
def api_terminals(
    q: str = Query("", description="Строка поиска (номер объекта, город, IP, VLAN...)"),
    tm: str = Query("", description="Тип билетомата: BPA или APB"),
    vendor: str = Query("", description="Производитель терминала"),
    scheme: str = Query("", description="Схема адресации: static / dhcp"),
    city: str = Query("", description="Город"),
    source: str = Query("", description="'ref' — справочник, 'obj' — объекты из файла"),
):
    res = filter_all(q, tm, vendor, scheme, city, source)
    return {"count": len(res), "results": res}


@app.post("/api/upload")
async def upload(file: UploadFile = File(...)):
    """Загрузить XLS/XLSX с объектной базой. Все листы импортируются автоматически."""
    global UPLOADED
    name = file.filename or ""
    if not re.search(r"\.(xls|xlsx|xlsm)$", name, re.I):
        return JSONResponse({"ok": False,
                             "error": "Нужен файл .xls или .xlsx"}, status_code=400)
    blob = await file.read()
    if len(blob) > 30 * 1024 * 1024:
        return JSONResponse({"ok": False, "error": "Файл больше 30 МБ"}, status_code=400)
    ext = ".xlsx" if re.search(r"\.(xlsx|xlsm)$", name, re.I) else ".xls"
    tmp = os.path.join(BASE_DIR, "tmp_upload" + ext)
    try:
        with open(tmp, "wb") as f:
            f.write(blob)
        records, sheets = xls_import.parse_file(tmp)
    except Exception as e:
        return JSONResponse({"ok": False, "error": f"Не удалось прочитать файл: {e}"},
                            status_code=400)
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)
    if not records:
        return JSONResponse({"ok": False,
                             "error": "В файле не найдено ни одной строки с настройками. "
                                      "Проверьте заголовки столбцов (город, номер, ip…)."},
                            status_code=400)
    cards = [xls_import.to_card(r, i + 1) for i, r in enumerate(records)]
    UPLOADED = cards
    _save_uploaded(cards)
    # из имени файла убираем бренд/ведомственные пометки (СБЕР, Сбер и т.п.)
    disp_name = re.sub(r"(?i)(сбер|sber(?:bank)?|sberrf)", "", name)
    disp_name = re.sub(r"^[\s\-–_:.]+|[\s\-–_:.]+$", "", disp_name).strip() or "файл"
    meta = {"filename": disp_name, "uploaded_at": time.strftime("%Y-%m-%d %H:%M")}
    _save_meta(meta)
    cities = sorted({c["city"] for c in cards if c["city"]})
    return {"ok": True, "imported": len(cards), "meta": meta,
            "sheets": [{"name": n, "rows": k} for n, k in sheets],
            "cities": cities, "cards": cards}


@app.delete("/api/upload")
def clear_upload(_admin=Depends(require_admin)):
    global UPLOADED
    UPLOADED = []
    if os.path.exists(UPLOAD_STORE):
        os.remove(UPLOAD_STORE)
    if os.path.exists(META_FILE):
        os.remove(META_FILE)
    return {"ok": True}


@app.get("/api/state", response_class=JSONResponse)
def api_state():
    """Текущее состояние базы: файл, дата загрузки, количество объектов."""
    meta = _load_meta() if UPLOADED else {}
    return {"objects": len(UPLOADED), "meta": meta}


@app.get("/api/errors", response_class=JSONResponse)
def api_errors(q: str = Query("")):
    if not q:
        return {"count": len(ERRORS), "results": ERRORS}
    res = [e for e in ERRORS if _search(" ".join(e.values()), _norm(q))]
    return {"count": len(res), "results": res}


@app.get("/api/health")
def health():
    return {"status": "ok", "reference": len(TERMINALS), "objects": len(UPLOADED)}


# ------------------------------------------------------------- Web UI --------
with open(os.path.join(BASE_DIR, "page_ui.html"), encoding="utf-8") as _f:
    PAGE = _f.read()


@app.get("/", response_class=HTMLResponse)
def index():
    err_rows = ""
    html = PAGE.replace("__ERROR_ROWS__", err_rows)
    html = html.replace("__REF__", "[]")
    html = html.replace("__OBJ__", json.dumps(UPLOADED, ensure_ascii=False))
    return HTMLResponse(html, headers={
        "Cache-Control": "no-cache, no-store, must-revalidate",
        "Pragma": "no-cache", "Expires": "0",
    })


if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", "9000"))
    uvicorn.run(app, host="0.0.0.0", port=port)
