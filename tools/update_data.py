#!/usr/bin/env python3
"""Обновление данных об инфляции и ключевой ставке.

Что делает:
  * ключевая ставка — скачивает историю изменений ставки с сайта Банка России и
    считает наибольшую ставку на 1-е число месяцев каждого года (нужна для налога
    на проценты по вкладам) -> data/key_rate.json;
  * инфляция — добавляет новые месяцы в data/inflation.json из одного из источников:
      1) --from-xlsx файл     таблица Росстата «ИПЦ, месяцы (с 1991 г.)» (xlsx);
      2) --from-text файл     таблица, скопированная с сайта (строки «2026 1.62 0.73 …»);
      3) без параметров       скачивает xlsx по адресу из tools/sources.json.

Главный принцип — НЕ ПОРТИТЬ данные. Перед записью всё проверяется (см. validate_*):
если что-то подозрительно, скрипт ничего не меняет, печатает причину и завершается
с кодом 1 (в GitHub Actions это даёт письмо о сбое).

Запуск:  python tools/update_data.py [--dry-run] [--from-xlsx F | --from-text F]
Нужен только openpyxl (pip install openpyxl) и только для чтения xlsx.
"""
import argparse, json, os, re, sys, tempfile, urllib.request, xml.etree.ElementTree as ET
from datetime import date, datetime

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
INFL_FILE = os.path.join(ROOT, "data", "inflation.json")
KEY_FILE = os.path.join(ROOT, "data", "key_rate.json")
SRC_FILE = os.path.join(ROOT, "tools", "sources.json")

MAX_NEW_MONTHS = 3        # за один запуск не добавляем больше 3 новых месяцев
REVISION_TOLERANCE = 0.10 # старые месяцы могут отличаться не больше чем на 0,10 п.п. (иначе — стоп)
MONTH_MIN, MONTH_MAX = -10.0, 300.0   # допустимая месячная инфляция, % (1992 год — до 245 %)


class DataError(Exception):
    """Данные не прошли проверку — запись отменена."""


# ---------- чтение и запись JSON ----------

def load_json(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def save_json(path, data):
    """Атомарная запись: сначала во временный файл, потом замена — файл не останется «недописанным»."""
    fd, tmp = tempfile.mkstemp(dir=os.path.dirname(path), suffix=".tmp")
    with os.fdopen(fd, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)
    os.replace(tmp, path)


# ---------- разбор источников инфляции ----------

def compound(values):
    p = 1.0
    for v in values:
        p *= 1 + v / 100
    return (p - 1) * 100


def strip_total(vals):
    """В таблицах за последним месяцем года часто идёт столбец «Всего» (итог за год).
    Убираем его: он либо 13-й по счёту, либо равен сложению предыдущих значений."""
    if len(vals) == 13:
        return vals[:12]
    if len(vals) >= 2 and abs(compound(vals[:-1]) - vals[-1]) <= max(0.06, abs(vals[-1]) * 0.002):
        return vals[:-1]
    return vals


def parse_text_table(text):
    """Строки вида «2026 1.62 0.73 0.60 …» (как в таблице на сайте) -> {год: [значения]}."""
    result = {}
    for line in text.splitlines():
        parts = line.replace(",", ".").replace("−", "-").split()
        if len(parts) < 2 or not re.fullmatch(r"(19|20)\d\d", parts[0]):
            continue
        try:
            vals = [float(x) for x in parts[1:]]
        except ValueError:
            continue
        result[int(parts[0])] = strip_total(vals)
    return result


def parse_xlsx(path):
    """Ищет в книге строки «год + числа». Если все числа похожи на индексы (100,62),
    переводит их в проценты (0,62). Формат файла Росстата может меняться, поэтому
    результат всё равно проходит проверку validate_inflation."""
    import openpyxl
    result = {}
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    for ws in wb.worksheets:
        for row in ws.iter_rows(values_only=True):
            cells = [c for c in row if c not in (None, "")]
            if len(cells) < 2:
                continue
            first = cells[0]
            if isinstance(first, str) and re.fullmatch(r"\d{4}", first.strip()):
                first = int(first)
            if not (isinstance(first, (int, float)) and 1991 <= int(first) <= 2100):
                continue
            nums = [float(c) for c in cells[1:] if isinstance(c, (int, float))]
            if len(nums) < 1:
                continue
            nums = nums[:12]                                   # первые 12 чисел — месяцы
            if all(v > 50 for v in nums):                      # индекс к предыдущему месяцу: 100,62 -> 0,62
                nums = [round(v - 100, 2) for v in nums]
            result.setdefault(int(first), nums)                # первая найденная таблица на листе — основная
    if not result:
        raise DataError("В xlsx не найдено ни одной строки «год + значения по месяцам»")
    return result


def fetch(url, timeout=30):
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (data-updater)"})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return resp.read()


# ---------- проверка и слияние инфляции ----------

def validate_inflation(old, new):
    """old, new: {год(int): [месячные значения]}. Возвращает слитые данные и список новых месяцев."""
    merged, added = {y: list(v) for y, v in old.items()}, []
    last_old = (max(old), len(old[max(old)]))
    for year in sorted(new):
        vals = new[year]
        if not vals:
            continue
        for i, v in enumerate(vals):
            if not (MONTH_MIN <= v <= MONTH_MAX):
                raise DataError(f"{year}-{i+1:02d}: значение {v} вне допустимых пределов — похоже на сбой разбора")
        if year in old:
            for i, v in enumerate(vals[:len(old[year])]):          # сверяем уже известные месяцы
                if abs(v - old[year][i]) > REVISION_TOLERANCE:
                    raise DataError(f"{year}-{i+1:02d}: в источнике {v}, у нас {old[year][i]} — расхождение больше "
                                    f"{REVISION_TOLERANCE} п.п. (возможна ошибка разбора или пересмотр данных, проверьте вручную)")
            extra = vals[len(old[year]):]
        else:
            extra = vals
        if extra:
            if year < last_old[0] or (year > last_old[0] and len(old[last_old[0]]) < 12):
                raise DataError(f"Новые данные за {year} год, но предыдущий период не завершён — пропуск месяцев")
            merged.setdefault(year, [])
            merged[year] += extra
            added += [(year, len(merged[year]) - len(extra) + k + 1, x) for k, x in enumerate(extra)]
            if len(merged[year]) > 12:
                raise DataError(f"{year}: больше 12 месяцев")
    if len(added) > MAX_NEW_MONTHS:
        raise DataError(f"За один запуск найдено {len(added)} новых месяцев (больше {MAX_NEW_MONTHS}) — нужна ручная проверка")
    return merged, added


def update_inflation(new_table, dry):
    doc = load_json(INFL_FILE)
    old = {int(y): v for y, v in doc["monthly"].items()}
    merged, added = validate_inflation(old, new_table)
    if not added:
        print("Инфляция: новых месяцев нет.")
        return False
    for y, m, v in added:
        print(f"Инфляция: +{y}-{m:02d} = {v:+.2f} %")
    if not dry:
        last_y = max(merged)
        doc["monthly"] = {str(y): merged[y] for y in sorted(merged)}
        doc["last_month"] = f"{last_y}-{len(merged[last_y]):02d}"
        doc["updated"] = date.today().isoformat()
        save_json(INFL_FILE, doc)
    return True


# ---------- ключевая ставка ----------

def parse_key_rate_xml(xml_bytes):
    """Достаёт пары (дата, ставка) из XML Банка России. Названия тегов ищем гибко
    (в имени есть «dt»/«date» и «rate»), чтобы мелкие изменения формата не ломали разбор."""
    root, events = ET.fromstring(xml_bytes), []
    for el in root.iter():
        kids = {c.tag.split("}")[-1].lower(): (c.text or "").strip() for c in el}
        d = next((v for k, v in kids.items() if k in ("dt", "date") or k.endswith("date")), None)
        r = next((v for k, v in kids.items() if "rate" in k or k in ("val", "value")), None)
        if d and r:
            try:
                events.append((datetime.fromisoformat(d[:10]).date(), float(r.replace(",", "."))))
            except ValueError:
                continue
    if not events:
        raise DataError("В ответе Банка России не найдено ни одной ставки")
    return sorted(set(events))


def max_rates_by_year(events, today=None):
    """events — [(дата, ставка)] (даты записей; между записями ставка не меняется).
    Для каждого года — наибольшая ставка на 1-е число любого месяца (учитываем только
    уже наступившие месяцы). Также возвращает текущую ставку."""
    today = today or date.today()

    def rate_on(day):
        past = [r for d, r in events if d <= day]
        return past[-1] if past else None

    result = {}
    for year in range(2023, today.year + 1):
        firsts = [rate_on(date(year, m, 1)) for m in range(1, 13) if date(year, m, 1) <= today]
        firsts = [r for r in firsts if r is not None]
        if firsts:
            result[year] = max(firsts)
    return result, rate_on(today)


def update_key_rate(events, dry):
    doc = load_json(KEY_FILE)
    new, current = max_rates_by_year(events)
    old = {int(y): v for y, v in doc["max_rate_by_year"].items()}
    for y, v in new.items():
        if y < max(old) and abs(v - old[y]) > 1e-9:           # прошлые годы не должны меняться
            raise DataError(f"Ключевая ставка {y}: в источнике {v}, у нас {old[y]} — расхождение, запись отменена")
        if not (0 < v <= 100):
            raise DataError(f"Ключевая ставка {y}: значение {v} вне пределов")
    changed = {y: v for y, v in new.items() if old.get(y) != v} or ({} if current == doc.get("current_rate") else {0: current})
    if not changed:
        print("Ключевая ставка: без изменений.")
        return False
    print("Ключевая ставка:", {y: v for y, v in new.items() if old.get(y) != v}, "текущая:", current)
    if not dry:
        doc["max_rate_by_year"] = {str(y): new[y] for y in sorted(new)}
        doc["current_rate"] = current
        doc["updated"] = date.today().isoformat()
        save_json(KEY_FILE, doc)
    return True


# ---------- запуск ----------

def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--dry-run", action="store_true", help="только показать, что изменится")
    ap.add_argument("--from-xlsx"), ap.add_argument("--from-text")
    ap.add_argument("--skip-key-rate", action="store_true"), ap.add_argument("--skip-inflation", action="store_true")
    a = ap.parse_args()
    sources = load_json(SRC_FILE) if os.path.exists(SRC_FILE) else {}
    problems, changed = [], False

    if not a.skip_inflation:
        try:
            if a.from_text:
                table = parse_text_table(open(a.from_text, encoding="utf-8").read())
            elif a.from_xlsx:
                table = parse_xlsx(a.from_xlsx)
            elif sources.get("inflation_xlsx_url"):
                path = os.path.join(tempfile.gettempdir(), "cpi.xlsx")
                open(path, "wb").write(fetch(sources["inflation_xlsx_url"]))
                table = parse_xlsx(path)
            else:
                table = None
                print("Инфляция: источник не настроен — пропуск (укажите inflation_xlsx_url в tools/sources.json "
                      "или запускайте с --from-text / --from-xlsx).")
            if table is not None:
                changed |= update_inflation(table, a.dry_run)
        except (DataError, OSError, ValueError) as e:
            problems.append(f"Инфляция: {e}")

    if not a.skip_key_rate:
        try:
            url = sources.get("key_rate_url")
            if not url:
                print("Ключевая ставка: адрес не настроен — пропуск.")
            else:
                url = url.replace("{from}", "01/01/2022").replace("{to}", date.today().strftime("%m/%d/%Y"))
                changed |= update_key_rate(parse_key_rate_xml(fetch(url)), a.dry_run)
        except (DataError, OSError, ET.ParseError, ValueError) as e:
            problems.append(f"Ключевая ставка: {e}")

    for p in problems:
        print("ОШИБКА —", p, file=sys.stderr)
    print("Готово: данные " + ("обновлены." if changed and not a.dry_run else "не изменялись."))
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main())
