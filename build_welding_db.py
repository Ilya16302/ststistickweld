#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import argparse
import glob
import gzip
import hashlib
import json
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict
from datetime import datetime, timedelta
from pathlib import Path

NS_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
NS_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

DEFAULT_ADMIN_LOGIN = "admin"
DEFAULT_ADMIN_PASSWORD = "smk_engineer1103"

JOINT_FIELDS = [
    "id", "source_row", "isometry", "joint_type", "joint_no", "weld_date",
    "vt_request", "rt_request", "ut_request", "welding_method", "material",
    "diameter_mm", "thickness_mm", "production_value", "result", "foreman",
    "brigadier", "welder_text", "welder_ids", "ut_date", "ut_result",
    "rt_date", "rt_result", "defects", "defect_events"
]

DEFECT_DECODER = {
    "Ea": "трещина вдоль шва",
    "Eb": "трещина поперек шва",
    "Ec": "трещина разветвленная",
    "Da": "непровар в корне шва",
    "Db": "непровар между валиками шва",
    "Dc": "непровар по разделке шва",
    "Aa": "пора",
    "Ab": "цепочка пор",
    "Ac": "скопление пор",
    "Ba": "шлак",
    "Bb": "цепочка шлака",
    "Bc": "скопление шлака",
    "Ca": "вольфрам",
    "Cb": "цепочка вольфрама",
    "Cc": "скопление вольфрама",
    "Fa": "вогнутость корня шва",
    "Fb": "выпуклость корня шва",
    "Fc": "подрез",
    "Fd": "смещение кромок",
}


def col_to_num(col):
    n = 0
    for ch in col:
        n = n * 26 + ord(ch.upper()) - 64
    return n

def cell_ref_parts(ref):
    m = re.match(r"([A-Z]+)(\d+)", ref)
    return m.group(1), int(m.group(2))

def clean_val(v):
    if v is None:
        return None
    if isinstance(v, str):
        v = v.strip()
        if v == "":
            return None
        return v
    if isinstance(v, float) and v.is_integer():
        return int(v)
    return v

def norm_text(s):
    if s is None:
        return ""
    s = str(s).strip().lower().replace("ё", "е")
    s = re.sub(r"\s+", " ", s)
    return s

def sha256_text(s):
    return hashlib.sha256((s or "").encode("utf-8")).hexdigest()

def stable_id(*parts, length=12):
    base = "|".join("" if p is None else str(p) for p in parts)
    return hashlib.sha1(base.encode("utf-8")).hexdigest()[:length]

def excel_date_to_iso(v):
    v = clean_val(v)
    if v is None:
        return None
    if isinstance(v, (int, float)):
        try:
            d = datetime(1899, 12, 30) + timedelta(days=float(v))
            if 1900 <= d.year <= 2100:
                return d.date().isoformat()
        except Exception:
            pass
    if isinstance(v, str):
        s = v.strip()
        if not s or s == "-":
            return None
        for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y", "%d.%m.%y"):
            try:
                return datetime.strptime(s, fmt).date().isoformat()
            except ValueError:
                pass
    return str(v)

def load_shared_strings(z):
    try:
        with z.open("xl/sharedStrings.xml") as f:
            out = []
            for event, elem in ET.iterparse(f, events=("end",)):
                if elem.tag == NS_MAIN + "si":
                    out.append("".join((t.text or "") for t in elem.iter(NS_MAIN + "t")))
                    elem.clear()
            return out
    except KeyError:
        return []

def get_sheet_paths(z):
    ns = {
        "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    }
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    relmap = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}
    result = {}
    for sh in wb.find("main:sheets", ns):
        name = sh.attrib["name"]
        rid = sh.attrib[NS_REL + "id"]
        target = relmap[rid]
        result[name] = "xl/" + target if not target.startswith("xl/") else target
    return result

def parse_sheet_rows(z, sst, sheet_path, min_row=1, max_row=None, target_cols=None):
    target_nums = None
    if target_cols:
        target_nums = {col_to_num(c) if isinstance(c, str) else c for c in target_cols}

    with z.open(sheet_path) as f:
        for event, row in ET.iterparse(f, events=("end",)):
            if row.tag != NS_MAIN + "row":
                continue
            r_idx = int(row.attrib.get("r", "0"))
            if r_idx >= min_row and (max_row is None or r_idx <= max_row):
                vals = {}
                for c in row.findall(NS_MAIN + "c"):
                    ref = c.attrib.get("r")
                    if not ref:
                        continue
                    col, _ = cell_ref_parts(ref)
                    if target_nums is not None and col_to_num(col) not in target_nums:
                        continue
                    cell_type = c.attrib.get("t")
                    v_el = c.find(NS_MAIN + "v")
                    is_el = c.find(NS_MAIN + "is")
                    val = None

                    if cell_type == "s":
                        if v_el is not None and v_el.text is not None:
                            idx = int(v_el.text)
                            val = sst[idx] if 0 <= idx < len(sst) else ""
                    elif cell_type == "inlineStr":
                        if is_el is not None:
                            val = "".join((te.text or "") for te in is_el.iter(NS_MAIN + "t"))
                    elif cell_type == "b":
                        val = v_el is not None and v_el.text == "1"
                    else:
                        val = v_el.text if v_el is not None else None
                        if val is not None:
                            try:
                                if re.fullmatch(r"-?\d+", val):
                                    val = int(val)
                                else:
                                    val = float(val)
                            except Exception:
                                pass
                    vals[col] = val

                yield r_idx, vals

            row.clear()
            if max_row is not None and r_idx > max_row:
                break

def determine_result(bl, af, ai):
    bl = clean_val(bl)
    if bl:
        return str(bl)
    for v in (af, ai):
        v = clean_val(v)
        if isinstance(v, str) and v.strip().upper().startswith("GCC"):
            return "Ожидание контроля"
    return "Контроль не назначен"

def split_welder_text(text):
    if not text:
        return []
    parts = re.split(r"\s*(?:/|;|\n|\+|,|–|—|-)\s*", str(text))
    return [p.strip() for p in parts if p and p.strip() and p.strip() != "-"]


def decode_defects(value):
    """Расшифровать коды дефектов из Excel в читаемый текст.

    Сохраняем исходные размеры/координаты, но подставляем смысл кода:
    "Aa3x0,8(495-500)" -> "Aa — пора 3x0,8(495-500)".
    "3Fc5(590-610)" -> "3Fc — подрез 5(590-610)".
    Если известных кодов нет, возвращается исходный текст.
    """
    value = clean_val(value)
    if not value:
        return None

    text = str(value).strip()

    def repl(match):
        code = match.group(1)
        desc = DEFECT_DECODER.get(code)
        return f"{code} — {desc} " if desc else code

    decoded = re.sub(r"(?<![A-Za-z])([A-F][a-d])(?![a-z])", repl, text)
    decoded = re.sub(r"\s+([,;:)])", r"\1", decoded)
    decoded = re.sub(r"\(\s+", "(", decoded)
    decoded = re.sub(r"\s{2,}", " ", decoded).strip()
    return decoded


def no_rk_marker(value):
    value = clean_val(value)
    if not value:
        return False
    s = str(value)
    stripped = s.strip()
    return (
        s.startswith(" / НО РК ")
        or s.startswith("НО РК")
        or stripped.startswith("/ НО РК")
        or stripped.startswith("НО РК")
    )

def determine_defects(result, br, ao, aq, au, bg, bm, av):
    if result != "Не годен":
        return None

    if no_rk_marker(br):
        return None

    candidates = []
    for col, value in (("AO", ao), ("AQ", aq), ("AU", au), ("BG", bg)):
        d = excel_date_to_iso(value)
        if d:
            candidates.append((d, col))

    if not candidates:
        return None

    earliest_date, earliest_col = min(candidates, key=lambda item: (item[0], {"AO": 0, "AQ": 1, "AU": 2, "BG": 3}[item[1]]))

    if earliest_col in ("AO", "AQ"):
        return decode_defects(bm)
    if earliest_col == "AU":
        return decode_defects(av)
    if earliest_col == "BG":
        return "Вырез"
    return None

def add_named_event(container, kind, number, method, event_date, defect_location=None):
    method = clean_val(method)
    event_date = excel_date_to_iso(event_date)
    defect_location = clean_val(defect_location)
    if method or event_date or defect_location:
        item = {"kind": kind, "number": number, "method": method, "date": event_date}
        if defect_location:
            item["defect_location"] = defect_location
        container.append(item)

def parse_stats(z, sst, sheet_path):
    stats_cols = ["C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "X", "Y", "AE", "AF"]
    welders = []
    period_start = None
    period_end = None

    for row_num, vals in parse_sheet_rows(z, sst, sheet_path, min_row=3, target_cols=stats_cols):
        if row_num == 3:
            period_start = excel_date_to_iso(vals.get("X"))
            period_end = excel_date_to_iso(vals.get("Y"))
            continue

        name = clean_val(vals.get("G"))
        login = clean_val(vals.get("AE"))
        if not name and not login:
            continue

        password = clean_val(vals.get("AF")) or ""
        wid = stable_id("welder", row_num, norm_text(name), login, length=12)

        welders.append({
            "id": wid,
            "source_row": row_num,
            "shift": clean_val(vals.get("C")),
            "position": clean_val(vals.get("D")),
            "responsible": clean_val(vals.get("E")),
            "role": clean_val(vals.get("F")),
            "name": name,
            "stamp": clean_val(vals.get("H")),
            "status": clean_val(vals.get("I")),
            "organization": clean_val(vals.get("J")),
            "login": norm_text(login),
            "password_hash": sha256_text(password),
            "password_plain": str(password) if password is not None else "",
            "stats_all": {
                "total_welded": clean_val(vals.get("K")) or 0,
                "production": clean_val(vals.get("L")) or 0,
                "submitted_control": clean_val(vals.get("M")) or 0,
                "accepted": clean_val(vals.get("N")) or 0,
                "rejected": clean_val(vals.get("O")) or 0,
                "reject_rate": clean_val(vals.get("P")),
            },
            "stats_period": {
                "total_welded": clean_val(vals.get("Q")) or 0,
                "production": clean_val(vals.get("R")) or 0,
                "submitted_control": clean_val(vals.get("S")) or 0,
                "accepted": clean_val(vals.get("T")) or 0,
                "rejected": clean_val(vals.get("U")) or 0,
                "reject_rate": clean_val(vals.get("V")),
            },
        })

    return welders, period_start, period_end

def parse_joints(z, sst, sheet_path, welders):
    cols = [
        "B", "C", "E", "F", "H", "J", "S", "Y", "AA", "AB", "AC", "AF", "AI",
        "AO", "AP", "AQ", "AR", "AL", "AM", "AN", "AS", "BL", "BM", "BR",
        "AT", "AU", "AV", "AW", "AX", "AY", "AZ", "BA", "BB", "BC", "BD", "BE",
        "BF", "BG", "BH", "BI", "BJ", "BK"
    ]

    name_to_ids = defaultdict(list)
    for w in welders:
        n = norm_text(w.get("name"))
        if n:
            name_to_ids[n].append(w["id"])

    joints = []
    total_rows = 0
    included_rows = 0
    rows_with_welder = 0
    rows_matched_to_welder = 0

    for row_num, vals in parse_sheet_rows(z, sst, sheet_path, min_row=4, target_cols=cols):
        total_rows += 1
        weld_date = excel_date_to_iso(vals.get("H"))
        status = clean_val(vals.get("AS"))

        if not weld_date or str(status).strip().upper() != "I":
            continue

        included_rows += 1
        welder_text = clean_val(vals.get("AN"))

        welder_ids = []
        seen = set()
        if welder_text:
            rows_with_welder += 1
            for part in split_welder_text(welder_text):
                n = norm_text(part)
                for wid in name_to_ids.get(n, []):
                    if wid not in seen:
                        welder_ids.append(wid)
                        seen.add(wid)
            if welder_ids:
                rows_matched_to_welder += 1

        events = []
        add_named_event(events, "repair", 1, vals.get("AT"), vals.get("AU"), vals.get("AV"))
        add_named_event(events, "repair", 2, vals.get("AW"), vals.get("AX"), vals.get("AY"))
        add_named_event(events, "repair", 3, vals.get("AZ"), vals.get("BA"), vals.get("BB"))
        add_named_event(events, "repair", 4, vals.get("BC"), vals.get("BD"), vals.get("BE"))
        add_named_event(events, "cut", 1, vals.get("BF"), vals.get("BG"))
        add_named_event(events, "cut", 2, vals.get("BH"), vals.get("BI"))
        add_named_event(events, "cut", 3, vals.get("BJ"), vals.get("BK"))

        nps_1 = clean_val(vals.get("S"))
        nps_2 = clean_val(vals.get("Y"))
        diameter_mm = clean_val(vals.get("AA"))
        production_value = nps_2 or nps_1
        if not isinstance(production_value, (int, float)) and isinstance(diameter_mm, (int, float)):
            production_value = round(float(diameter_mm) / 25.4, 2)

        result = determine_result(vals.get("BL"), vals.get("AF"), vals.get("AI"))
        defects = determine_defects(result, vals.get("BR"), vals.get("AO"), vals.get("AQ"), vals.get("AU"), vals.get("BG"), vals.get("BM"), vals.get("AV"))

        joint = {
            "id": stable_id("joint", row_num, vals.get("C"), vals.get("F"), length=14),
            "source_row": row_num,
            "isometry": clean_val(vals.get("C")),
            "joint_type": clean_val(vals.get("E")),
            "joint_no": clean_val(vals.get("F")),
            "weld_date": weld_date,
            "vt_request": clean_val(vals.get("AC")),
            "rt_request": clean_val(vals.get("AF")),
            "ut_request": clean_val(vals.get("AI")),
            "welding_method": clean_val(vals.get("J")),
            "material": clean_val(vals.get("B")),
            "diameter_mm": diameter_mm,
            "thickness_mm": clean_val(vals.get("AB")),
            "production_value": production_value,
            "result": result,
            "result_raw": clean_val(vals.get("BL")),
            "foreman": clean_val(vals.get("AL")),
            "brigadier": clean_val(vals.get("AM")),
            "welder_text": welder_text,
            "welder_ids": welder_ids,
            "ut_date": excel_date_to_iso(vals.get("AO")),
            "ut_result": clean_val(vals.get("AP")),
            "rt_date": excel_date_to_iso(vals.get("AQ")),
            "rt_result": clean_val(vals.get("AR")),
            "defects": defects,
            "defect_events": events or None,
        }

        joints.append([joint.get(field) for field in JOINT_FIELDS])

    return joints, {
        "book_rows_seen": total_rows,
        "primary_rows_included": included_rows,
        "rows_with_welder": rows_with_welder,
        "rows_matched_to_welder": rows_matched_to_welder,
        "rows_unmatched_to_welder": max(0, rows_with_welder - rows_matched_to_welder),
    }

def resolve_input(pattern):
    matches = sorted(glob.glob(pattern))
    if matches:
        return Path(matches[-1])
    p = Path(pattern)
    if p.exists():
        return p
    raise FileNotFoundError(f"Файл не найден: {pattern}")

def build_database(input_path, output_path, admin_login=DEFAULT_ADMIN_LOGIN, admin_password=DEFAULT_ADMIN_PASSWORD):
    with zipfile.ZipFile(input_path) as z:
        sst = load_shared_strings(z)
        sheet_paths = get_sheet_paths(z)

        missing = [name for name in ("Книга", "Статистика") if name not in sheet_paths]
        if missing:
            raise RuntimeError("Не найдены листы: " + ", ".join(missing))

        welders, period_start, period_end = parse_stats(z, sst, sheet_paths["Статистика"])
        joints, counters = parse_joints(z, sst, sheet_paths["Книга"], welders)

    defects_idx = JOINT_FIELDS.index("defects")
    defects_nonempty = sum(1 for row in joints if row[defects_idx])

    db = {
        "meta": {
            "source_file": str(input_path),
            "generated_at": datetime.now().isoformat(timespec="seconds"),
            "period_start": period_start,
            "period_end": period_end,
            "primary_rule": "Лист 'Книга': есть дата сварки в H и статус AS = I",
            "result_rule": "BL если заполнен; иначе GCC* в AF/AI = Ожидание контроля; иначе Контроль не назначен",
            "defects_rule": "Для Не годен: если BR начинается с НО РК — дефекты не пишем; иначе самая ранняя дата из AO/AQ/AU/BG: AO/AQ -> BM, AU -> AV, BG -> Вырез. Коды дефектов расшифровываются по словарю DEFECT_DECODER.",
            "password_storage": "sha256(password) + password_plain для админ-прототипа; для боевого сайта убрать открытые пароли",
            "defects_nonempty": defects_nonempty,
            **counters,
        },
        "auth": {
            "admin_login": norm_text(admin_login),
            "admin_password_hash": sha256_text(admin_password),
        },
        "welders": welders,
        "joint_fields": JOINT_FIELDS,
        "joints": joints,
    }

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(json.dumps(db, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")

    # Для сайта удобнее держать рядом сжатую версию: index.html сначала ищет welding_db.json.gz.
    gz_path = output_path.with_suffix(output_path.suffix + ".gz")
    with gzip.open(gz_path, "wt", encoding="utf-8") as f:
        json.dump(db, f, ensure_ascii=False, separators=(",", ":"))

    return db

def main():
    parser = argparse.ArgumentParser(description="Собрать JSON-базу статистики сварщиков из xlsm файла.")
    parser.add_argument("input", nargs="?", default=r"\\10.36.0.251\gcc\СМУ\ОГС\Кузнецов И.С\Статистика\Статистика*.xlsm",
                        help="Путь к xlsm или маска, например \\\\server\\share\\Статистика*.xlsm")
    parser.add_argument("-o", "--output", default="welding_db.json", help="Куда сохранить JSON")
    args = parser.parse_args()

    input_path = resolve_input(args.input)
    db = build_database(input_path, Path(args.output))
    print("Готово.")
    print(f"Источник: {input_path}")
    output_path = Path(args.output).resolve()
    print(f"JSON: {output_path}")
    print(f"GZIP для сайта: {output_path.with_suffix(output_path.suffix + '.gz')}")
    print(f"Сварщиков: {len(db['welders'])}")
    print(f"Стыков первички: {len(db['joints'])}")
    print(f"Не сопоставлено по ФИО: {db['meta']['rows_unmatched_to_welder']}")

if __name__ == "__main__":
    main()
