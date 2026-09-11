# -*- coding: utf-8 -*-
"""
Обработка накладных:
  1) Витебские ковры (.xls) — см. VtebskieKovri_ДОКУМЕНТАЦИЯ.md
  2) Обои / УПД ВВП (.xlsx) — очистка наименований + коды из справочника
  3) Юсуф / packing list (.xlsx) — турецкие ковры, доп. колонки как в эталоне

Обои (УПД):
  Исходник поставщика содержит длинные наименования вида
    «588058 Victoria Stenova Revolute/Революция Обои винил … 1,06х10м, …»
  Эталон («… готовый .xlsx») — это УПД с:
    - кодом номенклатуры 1С (УТ… / УТ-…);
    - коротким наименованием коллекции (как в справочнике);
    - артикулом в отдельной колонке.
  Скрипт читает сырой УПД/.xlsx, достаёт артикул из начала строки,
  подставляет код и короткое имя из oboi_catalog.json (если артикул
  известен), иначе строит имя эвристикой. Результат —
  «<имя>_ОБОИ_Обработано.xlsx», лист «Обработанные данные».

Использование:
    python convert_nakladnaya.py
    python convert_nakladnaya.py <файл>
    python convert_nakladnaya.py <файл> <результат.xlsx>
    python convert_nakladnaya.py --force
    python convert_nakladnaya.py --learn-oboi "… готовый .xlsx"
    python convert_nakladnaya.py --setup-inbox
    python convert_nakladnaya.py --learn-inbox
    python convert_nakladnaya.py --watch
"""

import sys
import os
import re
import datetime
import json
import copy
import xlrd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment

# ---------------------------------------------------------------------------
# Общие настройки
# ---------------------------------------------------------------------------
STATE_FILENAME = ".vitebsk_processed.json"
OBOI_CATALOG_FILENAME = "oboi_catalog.json"
YUSUF_PROFILE_FILENAME = "yusuf_profile.json"

FONT_NAME = "Arial"

HEADER = [
    "НомерСтроки", "Тип", "Форма", "Наименование", "ПолноеНаименвоание",
    "ЕдиницаБазовая", "ЕдиницаХранения", "ЕдиницаХраненияКоэфф",
    "Характеристика", "Серия", "Ширина", "Длина", "Штрихкод",
    "Количество", "Сумма",
]

OBOI_HEADER = [
    "КодТовара", "НомерСтроки", "Наименование", "Артикул",
    "КодВидаТовара", "ЕдиницаКод", "Единица", "Количество",
    "Цена", "СуммаБезНДС", "Акциз", "СтавкаНДС", "СуммаНДС",
    "СуммаСНДС", "КодСтраны", "Страна", "НомерГТД",
]

VBA_DLINA_KOVROLIN_POROG = 5

_CYRILLIC_RE = re.compile(r"[а-яёА-ЯЁ]")
_OBOI_CODE_RE = re.compile(r"^УТ[-–]?\d", re.I)

# Бренды в длинном наименовании поставщика — их выкидываем при эвристике.
_OBOI_BRANDS = (
    "victoria stenova",
    "ateliero",
    "renowa",
)

# Известные правки имени, если артикула ещё нет в справочнике.
_OBOI_NAME_FIXES = {
    "лиза": "КЛЮКВА (Liza/Лиза)",
    "тренд": "ТРЕНД (6)",
    "bambi/бемби": "Bembi/Бемби (6)",
    "bambi / бемби": "Bembi/Бемби (6)",
}

# Пары, у которых в эталоне кириллица стоит первой.
_OBOI_REVERSE_PAIRS = {
    "dakar": "ДАКАР/ DAKAR (6)",
    "camelia": "КАМЕЛИЯ/ CAMELIA (6)",
    "mariinski": "МАРИИНСКИЙ/MARIINSKI (6)",
}


# ===========================================================================
# ЖУРНАЛ СООБЩЕНИЙ
# ===========================================================================
_log_sink = None


def set_log_sink(fn):
    """Перенаправляет сообщения скрипта (GUI подставляет сюда своё окно)."""
    global _log_sink
    _log_sink = fn


def log(msg="", level="info"):
    if _log_sink is not None:
        _log_sink(str(msg), level)
    else:
        print(msg)


# ===========================================================================
# ВСПОМОГАТЕЛЬНЫЕ (общие)
# ===========================================================================
def norm_num(v):
    """Число без лишнего .0, с запятой вместо точки — как в Excel RU."""
    if isinstance(v, float) and v.is_integer():
        s = str(int(v))
    else:
        s = str(v)
    return s.replace(".", ",")


def to_int_if_whole(v):
    if isinstance(v, float) and v.is_integer():
        return int(v)
    return v


def pervoe_slovo(text):
    s = str(text).strip()
    if not s:
        return ""
    return s.split()[0]


def app_dir():
    """Папка приложения: рядом с .exe при onefile-сборке, иначе рядом со скриптом."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


VERSION_FILENAME = "version.txt"
BUILT_IN_VERSION = "1.4"


def current_version():
    """Версия на диске (после «Обучить») или встроенная в сборку."""
    path = os.path.join(app_dir(), VERSION_FILENAME)
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                v = f.read().strip().splitlines()[0].strip()
            if v:
                return v
        except Exception:
            pass
    return BUILT_IN_VERSION


def save_version(version, folder=None):
    folder = folder or app_dir()
    path = os.path.join(folder, VERSION_FILENAME)
    with open(path, "w", encoding="utf-8") as f:
        f.write(str(version).strip() + "\n")
    return path


def bump_patch(version):
    parts = [p for p in str(version).split(".") if p != ""]
    nums = []
    for p in parts:
        try:
            nums.append(int(p))
        except ValueError:
            nums.append(0)
    while len(nums) < 3:
        nums.append(0)
    nums[-1] += 1
    return ".".join(str(n) for n in nums)


def bundle_dir():
    """Папка с ресурсами, вшитыми в onefile-сборку (справочник по умолчанию)."""
    return getattr(sys, "_MEIPASS", app_dir())


def script_dir():
    return app_dir()


def sanitize_dlya_imeni(text):
    bad = '<>:"/\\|?*'
    s = str(text)
    for ch in bad:
        s = s.replace(ch, "_")
    return s.replace(" ", "_")


def unik_put(path):
    """Если файл уже есть — суффикс (2), (3), …"""
    if not os.path.exists(path):
        return path
    base_path, ext = os.path.splitext(path)
    i = 2
    while os.path.exists(f"{base_path} ({i}){ext}"):
        i += 1
    return f"{base_path} ({i}){ext}"


# ===========================================================================
# ВИТЕБСК — определение коллекции / схемы
# ===========================================================================
def opredelit_kategoriya(produkciya):
    s = str(produkciya).lower().replace("ё", "е")
    if "ковер" in s:
        return "ковер"
    return "ковролин"


def opredelit_kollekciyu(risunok):
    s = str(risunok).strip()
    if len(s) > 1 and s[0] == "p" and s[1].isdigit():
        return "Палитра"
    if s.lower().startswith("sh/"):
        return "Шегги"
    parts = [p.strip() for p in s.split("/")]
    if len(parts) >= 3 and parts[2] and _CYRILLIC_RE.search(parts[2]):
        return parts[2].upper()
    return None


def transform_kover_imennoy(row):
    (num, produkciya, izdelie, risunok, shirina, dlina, sort, razbrak,
     kolvo, ed, ploshad, cena, summa, shtrih) = row[:14]

    parts = [p.strip() for p in str(risunok).split("/")]
    artikul = parts[0] if len(parts) > 0 else ""
    tsvet = parts[1] if len(parts) > 1 else ""
    kollekciya = parts[2] if len(parts) > 2 else ""

    tip = "Ковер"
    if "o/" in str(risunok) or "/ЭФ" in str(risunok):
        forma = "Овал"
    elif shirina == dlina:
        forma = "Круг"
    else:
        forma = "Прямоугольник"

    w_str = norm_num(shirina)
    l_str = norm_num(dlina)
    naimenovanie = f"{kollekciya} {w_str}*{l_str} {forma}".strip()
    polnoe = f"{tip} {naimenovanie}"
    harakteristika = f"{artikul} {tsvet}".strip()

    return [
        to_int_if_whole(num), tip, forma, naimenovanie, polnoe,
        "м2", ed, ploshad,
        harakteristika, "",
        shirina, dlina, shtrih, kolvo, summa,
    ]


def transform_kovrolin_imennoy(row, kollekciya):
    (num, produkciya, izdelie, risunok, shirina, dlina, sort, razbrak,
     kolvo, ed, ploshad, cena, summa, shtrih) = row[:14]

    parts = [p.strip() for p in str(risunok).split("/")]
    artikul = parts[0] if len(parts) > 0 else ""
    tsvet = parts[1] if len(parts) > 1 else ""

    tip = "Ковролин"
    forma = ""
    naimenovanie = kollekciya
    polnoe = f"{kollekciya} {pervoe_slovo(produkciya)}".strip()
    w_str = norm_num(shirina)
    harakteristika = f"{artikul} {tsvet} {w_str} м".strip()
    seria = to_int_if_whole(izdelie)

    return [
        to_int_if_whole(num), tip, forma, naimenovanie, polnoe,
        "м2", "м2", 1,
        harakteristika, seria,
        shirina, dlina, shtrih, ploshad, summa,
    ]


def ochistit_stroku_kovrolin(risunok):
    clean = str(risunok).strip()
    clean = clean.replace("/Шегги ", "").replace("Шегги ", "")
    clean = clean.replace("r/", "/").replace("p/", "/").replace("//", "/")
    if len(clean) > 1 and clean[0] == "p" and clean[1].isdigit():
        clean = clean[1:]
    return clean


def transform_kovrolin_ochistka(row):
    (num, produkciya, izdelie, risunok, shirina, dlina, sort, razbrak,
     kolvo, ed, ploshad, cena, summa, shtrih) = row[:14]

    tip = "Ковролин"
    forma = ""
    naimenovanie = ochistit_stroku_kovrolin(risunok)
    polnoe = f"{tip} {naimenovanie}"
    w_str = norm_num(shirina)
    harakteristika = f"{w_str} м"
    seria = to_int_if_whole(izdelie)

    return [
        to_int_if_whole(num), tip, forma, naimenovanie, polnoe,
        "м2", "м2", 1,
        harakteristika, seria,
        shirina, dlina, shtrih, ploshad, summa,
    ]


def transform_vba_obshiy(row):
    (num, produkciya, izdelie, risunok, shirina, dlina, sort, razbrak,
     kolvo, ed, ploshad, cena, summa, shtrih) = row[:14]

    pattern = str(risunok).strip()
    tip = "Ковролин" if dlina > VBA_DLINA_KOVROLIN_POROG else "Ковер"

    if tip == "Ковер":
        if "o/" in pattern or "/ЭФ" in pattern:
            forma = "Овал"
        elif shirina == dlina:
            forma = "Круг"
        else:
            forma = "Прямоугольник"
    else:
        forma = ""

    if tip == "Ковер":
        temp = pattern.replace("/Шегги ", "")
        if "//" in temp:
            parts = temp.split("//")
            if len(parts) >= 2:
                name_parts = parts[1].strip().split(" ")
                temp = name_parts[0] if name_parts and name_parts[0] else parts[1].strip()
        w_str = norm_num(shirina)
        l_str = norm_num(dlina)
        temp = f"{temp} {w_str}*{l_str} {forma}"
    else:
        temp = pattern.replace("Шегги ", "")

    temp = temp.replace("r/", "/").replace("p/", "/").replace("//", "/")
    naimenovanie = temp
    polnoe = f"{tip} {naimenovanie}"

    if tip == "Ковролин":
        w_str = norm_num(shirina)
        harakteristika = f"{w_str} м"
    else:
        harakteristika = ""
        if "//" in pattern:
            parts = pattern.split("//")
            if len(parts) >= 2:
                before = parts[0].strip()
                after = parts[1].strip()
                name_parts = after.split(" ")
                if len(name_parts) >= 2:
                    ostatok = after.replace(name_parts[0], "", 1).strip()
                    harakteristika = f"{before} {ostatok}".strip()
                else:
                    harakteristika = before

    ed_baz = "м2"
    if tip == "Ковролин":
        ed_hran = "м2"
        ed_koef = 1
        seria = to_int_if_whole(izdelie)
        kolichestvo = ploshad
    else:
        ed_hran = ed
        ed_koef = ploshad
        seria = ""
        kolichestvo = kolvo

    return [
        to_int_if_whole(num), tip, forma, naimenovanie, polnoe,
        ed_baz, ed_hran, ed_koef,
        harakteristika, seria,
        shirina, dlina, shtrih, kolichestvo, summa,
    ]


def transform_row(row):
    risunok = row[3]
    produkciya = row[1]
    kollekciya = opredelit_kollekciyu(risunok)

    if kollekciya in ("Палитра", "Шегги"):
        return transform_kovrolin_ochistka(row), kollekciya

    if kollekciya is not None:
        kat = opredelit_kategoriya(produkciya)
        if kat == "ковер":
            return transform_kover_imennoy(row), kollekciya
        return transform_kovrolin_imennoy(row, kollekciya), kollekciya

    return transform_vba_obshiy(row), None


def chitat_syrye_stroki(src_path):
    wb_src = xlrd.open_workbook(src_path)
    sh = wb_src.sheet_by_index(0)
    raw_rows = []
    for r in range(1, sh.nrows):
        raw = [sh.cell_value(r, c) for c in range(sh.ncols)]
        if raw[0] == "" or raw[0] is None:
            continue
        if not isinstance(raw[4], (int, float)) or not isinstance(raw[5], (int, float)):
            continue
        raw_rows.append(raw)
    return raw_rows


def postroit_imya_rezultata_vitebsk(src_path):
    folder = os.path.dirname(src_path) or "."
    base = os.path.splitext(os.path.basename(src_path))[0]
    return unik_put(os.path.join(folder, f"{base}_VITEBSK_Обработано.xlsx"))


def eto_gotovyj_vitebsk(path):
    try:
        wb = xlrd.open_workbook(path)
        return "Обработанные данные" in wb.sheet_names()
    except Exception:
        return True


def convert_vitebsk(src_path, out_path=None):
    raw_rows = chitat_syrye_stroki(src_path)
    kollekcii = []
    counts = {}
    processed = []

    for raw in raw_rows:
        out_row, kollekciya = transform_row(raw)
        processed.append(out_row)
        if kollekciya is not None and kollekciya not in kollekcii:
            kollekcii.append(kollekciya)
        label = kollekciya if kollekciya is not None else "не определена (общий VBA-алгоритм)"
        counts[label] = counts.get(label, 0) + 1

    if out_path is None:
        out_path = postroit_imya_rezultata_vitebsk(src_path)

    log(f"Файл: {os.path.basename(src_path)}  [Витебск]", "head")
    log("Обнаружено:")
    for label, cnt in sorted(counts.items(), key=lambda x: -x[1]):
        log(f"  - {label}: {cnt} стр.")
    if len(kollekcii) > 1:
        log("В файле несколько коллекций сразу — все они попадут в один")
        log(f"итоговый файл: {', '.join(kollekcii)}")
    log()

    wb_out = Workbook()
    ws = wb_out.active
    ws.title = "Обработанные данные"
    ws.append(HEADER)
    for cell in ws[1]:
        cell.font = Font(name=FONT_NAME, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for out_row in processed:
        ws.append(out_row)
        for cell in ws[ws.max_row]:
            cell.font = Font(name=FONT_NAME)
        ws.cell(row=ws.max_row, column=13).number_format = "0"

    widths = [10, 10, 14, 26, 32, 12, 14, 16, 14, 10, 9, 9, 16, 11, 10]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = w
    ws.freeze_panes = "A2"
    wb_out.save(out_path)
    return len(raw_rows), out_path


# ===========================================================================
# ЮСУФ — packing list / invoice ковров (.xlsx)
# Эталон: копируем лист поставщика и дописываем колонки 12–18
#   12–13 ширина/длина в см, 14 буква формы, 15 наим, 16 ед изм,
#   17 полн наим, 18 харка. Цвет без пробелов вокруг «/».
# ===========================================================================
_YUSUF_NAME_MARKERS = ("юсуф", "yusuf", "ysf", "packing")
YUSUF_SHAPE_DEFAULT = {
    "D": "Прямоугольник",
    "O": "Овал",
    "C": "Круг",
}
_YUSUF_TOTAL_MARKERS = ("TOTAL", "GENERAL TOTAL")


def default_yusuf_profile():
    return {
        "enabled": True,
        "unit": "шт",
        "name_prefix": "Ковер",
        "color_strip_slash_spaces": True,
        "meter_decimal": ",",
        "naim_template": "{collection} {w_m}*{l_m} {forma}",
        "full_template": "{prefix} {naim}",
        "char_template": "{design} {color}",
        "headers": {
            "naim": "наим",
            "unit": "ед изм",
            "full": "полн наим",
            "char": "харка",
        },
        "shape_letters": dict(YUSUF_SHAPE_DEFAULT),
    }


def yusuf_profile_path(folder=None):
    if folder is not None:
        return os.path.join(folder, YUSUF_PROFILE_FILENAME)
    external = os.path.join(app_dir(), YUSUF_PROFILE_FILENAME)
    if os.path.exists(external):
        return external
    bundled = os.path.join(bundle_dir(), YUSUF_PROFILE_FILENAME)
    if os.path.exists(bundled):
        return bundled
    return external


def load_yusuf_profile(folder=None):
    profile = default_yusuf_profile()
    path = yusuf_profile_path(folder)
    if not os.path.exists(path):
        return profile
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, dict):
            profile.update({k: v for k, v in data.items() if k != "shape_letters"})
            letters = data.get("shape_letters") or {}
            if isinstance(letters, dict):
                merged = dict(YUSUF_SHAPE_DEFAULT)
                merged.update({str(k).upper(): v for k, v in letters.items()})
                profile["shape_letters"] = merged
            headers = data.get("headers") or {}
            if isinstance(headers, dict):
                h = dict(default_yusuf_profile()["headers"])
                h.update(headers)
                profile["headers"] = h
    except Exception:
        pass
    return profile


def save_yusuf_profile(profile, folder=None):
    folder = folder or app_dir()
    path = os.path.join(folder, YUSUF_PROFILE_FILENAME)
    out = default_yusuf_profile()
    if isinstance(profile, dict):
        out.update(profile)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
        f.write("\n")
    return path


def _yusuf_v_imeni(path):
    name = os.path.basename(path).lower()
    return any(m in name for m in _YUSUF_NAME_MARKERS)


def _yusuf_chislo(v):
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(",", ".").strip())
    except Exception:
        return None


def yusuf_sm_kak_chislo(v):
    """Ширина/длина для колонок 12–13: как в эталоне, число в сантиметрах."""
    n = _yusuf_chislo(v)
    if n is None:
        return None
    if abs(n - round(n)) < 1e-9:
        return int(round(n))
    return n


def yusuf_sm_v_metry_str(v, profile=None):
    """Сантиметры → метры для наименования: 200→2, 150→1,5, 80→0,8."""
    n = _yusuf_chislo(v)
    if n is None:
        return ""
    if n >= 10:
        n = n / 100.0
    if abs(n - round(n)) < 1e-9:
        n = int(round(n))
    s = norm_num(n)
    dec = (profile or {}).get("meter_decimal", ",")
    if dec == ".":
        s = s.replace(",", ".")
    return s


def _yusuf_podstavit(tpl, mapping):
    out = str(tpl or "")
    for k in sorted(mapping, key=lambda x: -len(x)):
        out = out.replace("{" + k + "}", str(mapping[k]))
    return out


def yusuf_bukva_formy(size_shape):
    s = str(size_shape or "").strip().upper()
    if not s:
        return ""
    parts = s.replace("×", "x").replace("*", " x ").split()
    if parts and len(parts[-1]) == 1 and parts[-1].isalpha():
        return parts[-1]
    if "OVAL" in s or s.endswith(" O"):
        return "O"
    return ""


def yusuf_cvet(color, profile=None):
    s = str(color or "").strip()
    if profile is None or profile.get("color_strip_slash_spaces", True):
        s = re.sub(r"\s*/\s*", "/", s)
    return s


def yusuf_forma_ru(letter, width=None, length=None, profile=None):
    letters = (profile or default_yusuf_profile()).get("shape_letters") or {}
    key = str(letter or "").upper()
    if key and key in letters:
        return letters[key]
    w = _yusuf_chislo(width)
    l = _yusuf_chislo(length)
    if w is not None and l is not None and abs(w - l) < 1e-9:
        return letters.get("C") or "Круг"
    return letters.get("D") or "Прямоугольник"


def yusuf_dop_polya(raw, profile=None):
    """Дополнительные колонки эталона Юсуфа."""
    profile = profile or default_yusuf_profile()
    letter = yusuf_bukva_formy(raw.get("size_shape"))
    forma = yusuf_forma_ru(letter, raw.get("width"), raw.get("length"), profile)
    color = yusuf_cvet(raw.get("color"), profile)
    coll = str(raw.get("collection") or "").strip()
    design = str(raw.get("design") or "").strip()
    w_cm = yusuf_sm_kak_chislo(raw.get("width"))
    l_cm = yusuf_sm_kak_chislo(raw.get("length"))
    prefix = profile.get("name_prefix") or "Ковер"
    unit = profile.get("unit") or "шт"
    mapping = {
        "collection": coll,
        "design": design,
        "color": color,
        "w_m": yusuf_sm_v_metry_str(raw.get("width"), profile),
        "l_m": yusuf_sm_v_metry_str(raw.get("length"), profile),
        "w_cm": "" if w_cm is None else str(w_cm),
        "l_cm": "" if l_cm is None else str(l_cm),
        "forma": forma,
        "letter": letter or "D",
        "prefix": prefix,
        "unit": unit,
    }
    naim_tpl = profile.get("naim_template") or "{collection} {w_m}*{l_m} {forma}"
    naim = _yusuf_podstavit(naim_tpl, mapping).strip()
    mapping["naim"] = naim
    full_tpl = profile.get("full_template") or "{prefix} {naim}"
    char_tpl = profile.get("char_template") or "{design} {color}"
    return {
        "width_cm": w_cm,
        "length_cm": l_cm,
        "letter": letter or "D",
        "naim": naim,
        "unit": unit,
        "full": _yusuf_podstavit(full_tpl, mapping).strip(),
        "char": _yusuf_podstavit(char_tpl, mapping).strip(),
        "color": color,
        "forma": forma,
    }


def _yusuf_tekst_shapki(path, max_rows=35):
    try:
        wb = load_workbook(path, data_only=True, read_only=True)
        ws = wb[wb.sheetnames[0]]
        blob = []
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            if i >= max_rows:
                break
            for v in row:
                if v is not None:
                    blob.append(str(v).lower())
        wb.close()
        return " ".join(blob)
    except Exception:
        return ""


def eto_packing_list(path):
    text = _yusuf_tekst_shapki(path)
    return "packing list" in text and "collection name" in text


def eto_yusuf_gotovyj(path):
    """Уже готовый packing list Юсуфа (есть доп. колонки или слово «готовый»)."""
    if not str(path).lower().endswith(".xlsx"):
        return False
    bn = os.path.basename(path).lower()
    if "_юсуф_обработано" in bn:
        return True
    yusuf_name = _yusuf_v_imeni(path)
    if "готов" in bn and yusuf_name:
        return True
    if "_обработано" in bn and yusuf_name:
        return True
    try:
        wb = load_workbook(path, data_only=True, read_only=True)
        ws = wb[wb.sheetnames[0]]
        hits_naim = 0
        hits_packing = 0
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            if i > 40:
                break
            vals = [str(v).strip().lower() for v in row if v is not None]
            joined = " ".join(vals)
            if "packing list" in joined:
                hits_packing += 1
            if "collection name" in joined:
                hits_packing += 1
            if "наим" in vals or "полн наим" in vals or "харка" in vals:
                hits_naim += 1
        wb.close()
        if hits_naim and hits_packing:
            return True
        if "готов" in bn and hits_packing >= 2:
            return True
    except Exception:
        return False
    return False


def eto_yusuf_syroj(path):
    """Турецкий packing list (Юсуф): COLLECTION NAME / PACKING LIST / имя файла."""
    if not str(path).lower().endswith(".xlsx"):
        return False
    bn = os.path.basename(path).lower()
    if "_обработано" in bn:
        return False
    if eto_yusuf_gotovyj(path):
        return False
    if _yusuf_v_imeni(path):
        return True
    return eto_packing_list(path)


def nayti_yusuf_shapku(ws):
    header_row = None
    for r in range(1, min(ws.max_row, 40) + 1):
        vals = [str(ws.cell(r, c).value or "").strip().lower() for c in range(1, 16)]
        joined = " ".join(vals)
        if "collection" in joined and "quantity" in joined:
            header_row = r
            break
    if header_row is None:
        raise ValueError("Не найдена шапка packing list (COLLECTION NAME / Quantity)")

    cols = {
        "npp": 1, "collection": 2, "design": 3, "color": 4, "code": 5,
        "ean": 6, "width": 7, "length": 8, "m2_unit": 9, "qty": 10,
        "qty_m2": 11, "size_shape": 12,
    }
    for c in range(1, 16):
        title = str(ws.cell(header_row, c).value or "").strip().lower()
        title = re.sub(r"\s+", " ", title)
        if title in ("№", "no", "nо", "#"):
            cols["npp"] = c
        elif "collection" in title:
            cols["collection"] = c
        elif title == "design":
            cols["design"] = c
        elif title in ("color", "colour"):
            cols["color"] = c
        elif title == "code":
            cols["code"] = c
        elif "ean" in title:
            cols["ean"] = c
        elif "m2" in title and "unit" in title:
            cols["m2_unit"] = c
        elif title == "quantity":
            cols["qty"] = c
        elif "quantity" in title and "m2" in title:
            cols["qty_m2"] = c
        elif "size" in title and ("shope" in title or "shape" in title):
            cols["size_shape"] = c
        elif title == "the size":
            cols["width"] = c
            cols["length"] = c + 1

    sub_row = header_row + 1
    if str(ws.cell(sub_row, cols["width"]).value or "").strip().lower() == "width":
        data_start = sub_row + 1
    else:
        data_start = header_row + 1
        sub_row = None
    return header_row, sub_row, data_start, cols


def _yusuf_eto_itog(npp):
    s = str(npp or "").strip().upper()
    return s in _YUSUF_TOTAL_MARKERS or s.startswith("GENERAL TOTAL")


def chitat_yusuf_stroki(src_path):
    wb = load_workbook(src_path, data_only=True)
    ws = wb[wb.sheetnames[0]]
    header_row, sub_row, data_start, cols = nayti_yusuf_shapku(ws)
    rows = []
    first_total = None
    for r in range(data_start, ws.max_row + 1):
        npp = ws.cell(r, cols["npp"]).value
        if _yusuf_eto_itog(npp):
            first_total = r
            break
        if npp is None or str(npp).strip() == "":
            continue
        try:
            npp_n = int(float(npp))
        except Exception:
            continue
        coll = str(ws.cell(r, cols["collection"]).value or "").strip()
        qty = ws.cell(r, cols["qty"]).value
        if not coll or qty is None:
            continue
        rows.append({
            "row": r,
            "npp": npp_n,
            "collection": coll,
            "design": str(ws.cell(r, cols["design"]).value or "").strip(),
            "color": str(ws.cell(r, cols["color"]).value or "").strip(),
            "code": ws.cell(r, cols["code"]).value,
            "ean": ws.cell(r, cols["ean"]).value,
            "width": ws.cell(r, cols["width"]).value,
            "length": ws.cell(r, cols["length"]).value,
            "m2_unit": ws.cell(r, cols["m2_unit"]).value,
            "qty": qty,
            "size_shape": ws.cell(r, cols["size_shape"]).value,
        })
    wb.close()
    return rows, header_row, sub_row, data_start, first_total, cols


def _yusuf_ochistit_itogi(ws, first_total_row):
    if not first_total_row:
        return
    to_unmerge = []
    for mr in list(ws.merged_cells.ranges):
        if mr.min_row >= first_total_row:
            to_unmerge.append(str(mr))
    for addr in to_unmerge:
        ws.unmerge_cells(addr)
    max_col = max(ws.max_column, 18)
    for r in range(first_total_row, ws.max_row + 1):
        for c in range(1, max_col + 1):
            ws.cell(r, c).value = None


def postroit_imya_rezultata_yusuf(src_path):
    folder = os.path.dirname(src_path) or "."
    base = os.path.splitext(os.path.basename(src_path))[0]
    return unik_put(os.path.join(folder, f"{base}_ЮСУФ_Обработано.xlsx"))


def convert_yusuf(src_path, out_path=None, profile=None):
    """Копирует packing list и заполняет колонки эталона «готовый Юсуф»."""
    profile = profile or load_yusuf_profile()
    raw_rows, header_row, sub_row, data_start, first_total, cols = chitat_yusuf_stroki(src_path)
    if not raw_rows:
        raise ValueError("В packing list нет строк с товаром")
    if out_path is None:
        out_path = postroit_imya_rezultata_yusuf(src_path)

    counts = {}
    for r in raw_rows:
        counts[r["collection"]] = counts.get(r["collection"], 0) + 1

    log(f"Файл: {os.path.basename(src_path)}  [Юсуф / ковры]", "head")
    log(f"Обнаружено строк: {len(raw_rows)}")
    for label, cnt in sorted(counts.items(), key=lambda x: -x[1]):
        log(f"  - {label}: {cnt} стр.")
    log()

    wb = load_workbook(src_path)
    ws = wb[wb.sheetnames[0]]
    headers = profile.get("headers") or default_yusuf_profile()["headers"]
    label_row = sub_row or header_row
    ws.cell(label_row, 15).value = headers.get("naim", "наим")
    ws.cell(label_row, 16).value = headers.get("unit", "ед изм")
    ws.cell(label_row, 17).value = headers.get("full", "полн наим")
    ws.cell(label_row, 18).value = headers.get("char", "харка")
    sample_font = ws.cell(data_start, cols["collection"]).font
    header_font = ws.cell(label_row, cols["width"]).font
    for c in (15, 16, 17, 18):
        if header_font and header_font.name:
            ws.cell(label_row, c).font = copy.copy(header_font)
        else:
            ws.cell(label_row, c).font = Font(name="Times New Roman", bold=True, size=11)

    for raw in raw_rows:
        extra = yusuf_dop_polya(raw, profile)
        r = raw["row"]
        color_cell = ws.cell(r, cols["color"])
        color_cell.value = extra["color"]
        ws.cell(r, 12).value = extra["width_cm"]
        ws.cell(r, 13).value = extra["length_cm"]
        ws.cell(r, 14).value = extra["letter"]
        ws.cell(r, 15).value = extra["naim"]
        ws.cell(r, 16).value = extra["unit"]
        ws.cell(r, 17).value = extra["full"]
        ws.cell(r, 18).value = extra["char"]
        for c in range(12, 19):
            cell = ws.cell(r, c)
            if sample_font and sample_font.name:
                cell.font = copy.copy(sample_font)

    _yusuf_ochistit_itogi(ws, first_total)
    for letter, width in (("M", 8), ("N", 6), ("O", 32), ("P", 8), ("Q", 38), ("R", 32)):
        if not ws.column_dimensions[letter].width:
            ws.column_dimensions[letter].width = width
        elif letter in ("O", "Q", "R") and ws.column_dimensions[letter].width < 20:
            ws.column_dimensions[letter].width = width

    os.makedirs(os.path.dirname(os.path.abspath(out_path)) or ".", exist_ok=True)
    wb.save(out_path)
    return len(raw_rows), out_path


def sravnit_yusuf(out_path, got_path):
    """Сверка доп. колонок результата с эталоном. Список расхождений."""
    out_wb = load_workbook(out_path, data_only=True)
    got_wb = load_workbook(got_path, data_only=True)
    out_ws = out_wb[out_wb.sheetnames[0]]
    got_ws = got_wb[got_wb.sheetnames[0]]

    def data_rows(ws):
        header_row, sub_row, data_start, cols = nayti_yusuf_shapku(ws)
        rows = []
        for r in range(data_start, ws.max_row + 1):
            npp = ws.cell(r, cols["npp"]).value
            if _yusuf_eto_itog(npp):
                break
            try:
                int(float(npp))
            except Exception:
                continue
            coll = str(ws.cell(r, cols["collection"]).value or "").strip()
            if not coll:
                continue
            rows.append({
                "npp": to_int_if_whole(npp),
                "coll": coll,
                "design": str(ws.cell(r, cols["design"]).value or "").strip(),
                "color": yusuf_cvet(ws.cell(r, 4).value),
                "w": yusuf_sm_kak_chislo(ws.cell(r, 12).value),
                "l": yusuf_sm_kak_chislo(ws.cell(r, 13).value),
                "letter": str(ws.cell(r, 14).value or "").strip(),
                "naim": str(ws.cell(r, 15).value or "").strip(),
                "unit": str(ws.cell(r, 16).value or "").strip(),
                "full": str(ws.cell(r, 17).value or "").strip(),
                "char": str(ws.cell(r, 18).value or "").strip(),
            })
        return rows

    out_rows = data_rows(out_ws)
    got_rows = data_rows(got_ws)
    out_wb.close()
    got_wb.close()
    diffs = []
    if len(out_rows) != len(got_rows):
        diffs.append("разное число строк: результат %s, эталон %s" % (
            len(out_rows), len(got_rows)))
        return diffs
    fields = ["color", "w", "l", "letter", "naim", "unit", "full", "char"]
    for i, (o, g) in enumerate(zip(out_rows, got_rows), 1):
        bad = [f for f in fields if o[f] != g[f]]
        if bad:
            diffs.append("#%s: " % i + "; ".join(
                "%s: %r != %r" % (f, o[f], g[f]) for f in bad))
    return diffs


def _yusuf_chastoe(items, default):
    counts = {}
    for it in items:
        if not it:
            continue
        counts[it] = counts.get(it, 0) + 1
    if not counts:
        return default
    return sorted(counts.items(), key=lambda x: (-x[1], x[0]))[0][0]


def _yusuf_shablon_naim(naim, coll, w_m, l_m, forma):
    s = str(naim or "")
    if coll and s.startswith(coll):
        s = "{collection}" + s[len(coll):]
    if forma and s.endswith(forma):
        s = s[:-len(forma)] + "{forma}"
    if w_m and w_m in s:
        s = s.replace(w_m, "{w_m}", 1)
    if l_m and l_m in s:
        s = s.replace(l_m, "{l_m}", 1)
    return s


def _yusuf_shablon_prosto(text, mapping):
    s = str(text or "")
    items = [(len(str(v)), k, str(v)) for k, v in mapping.items() if str(v)]
    items.sort(reverse=True)
    for _, k, v in items:
        if v and v in s:
            s = s.replace(v, "{" + k + "}", 1)
    return s


def chitat_yusuf_gotovye_polya(got_path):
    wb = load_workbook(got_path, data_only=True)
    ws = wb[wb.sheetnames[0]]
    header_row, sub_row, data_start, cols = nayti_yusuf_shapku(ws)
    rows = []
    for r in range(data_start, ws.max_row + 1):
        npp = ws.cell(r, cols["npp"]).value
        if _yusuf_eto_itog(npp):
            break
        try:
            npp_n = int(float(npp))
        except Exception:
            continue
        coll = str(ws.cell(r, cols["collection"]).value or "").strip()
        if not coll:
            continue
        rows.append({
            "row": r,
            "npp": npp_n,
            "collection": coll,
            "design": str(ws.cell(r, cols["design"]).value or "").strip(),
            "color": str(ws.cell(r, 4).value or "").strip(),
            "width": ws.cell(r, cols["width"]).value,
            "length": ws.cell(r, cols["length"]).value,
            "width_cm": ws.cell(r, 12).value,
            "length_cm": ws.cell(r, 13).value,
            "letter": str(ws.cell(r, 14).value or "").strip(),
            "naim": str(ws.cell(r, 15).value or "").strip(),
            "unit": str(ws.cell(r, 16).value or "").strip(),
            "full": str(ws.cell(r, 17).value or "").strip(),
            "char": str(ws.cell(r, 18).value or "").strip(),
        })
    wb.close()
    return rows, header_row, sub_row


def learn_yusuf_from_gotovyj(got_path, folder=None, source_path=None):
    """Полностью перезаписывает правила Юсуфа по эталону (и исходнику, если есть)."""
    profile = default_yusuf_profile()
    got_rows, header_row, sub_row = chitat_yusuf_gotovye_polya(got_path)
    wb = load_workbook(got_path, data_only=True)
    ws = wb[wb.sheetnames[0]]
    label_row = sub_row or header_row
    headers = dict(default_yusuf_profile()["headers"])
    for key, col in (("naim", 15), ("unit", 16), ("full", 17), ("char", 18)):
        val = ws.cell(label_row, col).value
        if val:
            headers[key] = str(val).strip()
    wb.close()
    profile["headers"] = headers

    letters = dict(YUSUF_SHAPE_DEFAULT)
    units = []
    prefixes = []
    naim_tpls = []
    full_tpls = []
    char_tpls = []
    strip_slash = []
    decimals = []

    src_by_row = {}
    src_by_key = {}
    if source_path and os.path.exists(source_path):
        src_rows = chitat_yusuf_stroki(source_path)[0]
        for raw in src_rows:
            src_by_row[raw["row"]] = raw
            src_by_key[(raw["collection"], raw["npp"], raw["design"])] = raw

    n_rows = 0
    for got in got_rows:
        if not got["naim"]:
            continue
        n_rows += 1
        letter = (got["letter"] or "").upper()
        forma = got["naim"].split()[-1] if got["naim"].split() else ""
        if letter and forma:
            letters[letter] = forma
        if got["unit"]:
            units.append(got["unit"])
        raw = src_by_row.get(got["row"]) or src_by_key.get(
            (got["collection"], got["npp"], got["design"]))
        src_color = str(raw["color"]).strip() if raw else got["color"]
        stripped = re.sub(r"\s*/\s*", "/", src_color)
        if "/" in src_color:
            if got["color"] == stripped and got["color"] != src_color:
                strip_slash.append(True)
            elif got["color"] == src_color and src_color != stripped:
                strip_slash.append(False)

        width = raw["width"] if raw else got["width"]
        length = raw["length"] if raw else got["length"]
        trial = dict(profile)
        trial["shape_letters"] = letters
        w_m = yusuf_sm_v_metry_str(width, trial)
        l_m = yusuf_sm_v_metry_str(length, trial)
        w_dot = w_m.replace(",", ".")
        if w_m and w_m in got["naim"] and "," in w_m:
            decimals.append(",")
        elif w_dot and w_dot in got["naim"] and "." in w_dot:
            decimals.append(".")

        naim_tpls.append(_yusuf_shablon_naim(
            got["naim"], got["collection"], w_m, l_m, forma))
        if got["full"] and got["naim"] and got["full"].endswith(got["naim"]):
            prefix = got["full"][:-len(got["naim"])].strip()
            if prefix:
                prefixes.append(prefix)
                full_tpls.append("{prefix} {naim}" if got["full"] == prefix + " " + got["naim"]
                                 else _yusuf_shablon_prosto(got["full"], {
                                     "prefix": prefix, "naim": got["naim"]}))
            else:
                full_tpls.append("{naim}")
        color_norm = yusuf_cvet(src_color, {"color_strip_slash_spaces": True})
        char_tpls.append(_yusuf_shablon_prosto(got["char"], {
            "design": got["design"],
            "color": got["color"] or color_norm,
        }))

    profile["shape_letters"] = letters
    if units:
        profile["unit"] = _yusuf_chastoe(units, profile["unit"])
    if prefixes:
        profile["name_prefix"] = _yusuf_chastoe(prefixes, profile["name_prefix"])
    if naim_tpls:
        profile["naim_template"] = _yusuf_chastoe(
            naim_tpls, profile["naim_template"])
    if full_tpls:
        profile["full_template"] = _yusuf_chastoe(
            full_tpls, profile["full_template"])
    if char_tpls:
        profile["char_template"] = _yusuf_chastoe(
            char_tpls, profile["char_template"])
    if decimals:
        profile["meter_decimal"] = _yusuf_chastoe(decimals, ",")
    if strip_slash:
        profile["color_strip_slash_spaces"] = sum(1 for x in strip_slash if x) >= (
            len(strip_slash) / 2.0)

    profile["enabled"] = True
    profile["learned_from"] = os.path.basename(got_path)
    if source_path:
        profile["learned_source"] = os.path.basename(source_path)
    profile["learned_rows"] = n_rows
    profile["learned_at"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    path = save_yusuf_profile(profile, folder)
    log("Правила Юсуфа перезаписаны: %s" % path, "ok")
    log("  наим: %s" % profile.get("naim_template"))
    log("  полн: %s" % profile.get("full_template"))
    log("  харка: %s" % profile.get("char_template"))
    log("  ед.изм: %s; формы: %s" % (
        profile.get("unit"),
        ", ".join("%s→%s" % (k, v) for k, v in sorted(letters.items()))))
    return profile, path, n_rows


# ===========================================================================
# ОБОИ (УПД)
# ===========================================================================
def oboi_catalog_path(folder=None):
    """Свой справочник рядом с приложением важнее вшитого в сборку."""
    if folder is not None:
        return os.path.join(folder, OBOI_CATALOG_FILENAME)
    external = os.path.join(app_dir(), OBOI_CATALOG_FILENAME)
    if os.path.exists(external):
        return external
    return os.path.join(bundle_dir(), OBOI_CATALOG_FILENAME)


def load_oboi_catalog(folder=None):
    path = oboi_catalog_path(folder)
    if not os.path.exists(path):
        return {}
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_oboi_catalog(catalog, folder=None):
    folder = folder or app_dir()
    path = os.path.join(folder, OBOI_CATALOG_FILENAME)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(catalog, f, ensure_ascii=False, indent=2)


def artikul_as_key(v):
    """Артикул как строка-ключ справочника (без .0 у целых float)."""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    if isinstance(v, int):
        return str(v)
    s = str(v).strip()
    if re.fullmatch(r"\d+\.0+", s):
        return s.split(".", 1)[0]
    return s


def extract_artikul_i_hvost(full_name):
    """Первый токен — артикул, остальное — хвост описания."""
    s = str(full_name).strip()
    m = re.match(r"^(\S+)\s+(.*)$", s, flags=re.S)
    if not m:
        return s, ""
    return m.group(1).strip(), m.group(2).strip()


def oboi_heuristika_imeni(full_name):
    """Эвристика короткого имени, если артикула нет в справочнике."""
    _artikul, rest = extract_artikul_i_hvost(full_name)
    text = rest
    low = text.lower()
    for brand in _OBOI_BRANDS:
        if low.startswith(brand):
            text = text[len(brand):].strip()
            low = text.lower()
            break

    # Всё до слова «Обои»
    parts = re.split(r"\s+Обои\b", text, maxsplit=1, flags=re.I)
    coll = re.sub(r"\s+", " ", parts[0].strip())
    key = coll.lower().replace("ё", "е")
    if key in _OBOI_NAME_FIXES:
        return _OBOI_NAME_FIXES[key]

    if "/" in coll:
        left, right = [p.strip() for p in coll.split("/", 1)]
        left_key = left.lower()
        if left_key in _OBOI_REVERSE_PAIRS:
            return _OBOI_REVERSE_PAIRS[left_key]
        # Сохраняем пробелы вокруг «/», если они были в исходнике
        mid = coll[len(coll.split("/", 1)[0]):len(coll) - len(coll.split("/", 1)[1])]
        base = f"{left} / {right}" if " " in mid else f"{left}/{right}"
        if "(6)" not in base:
            base = f"{base} (6)"
        return base

    base = "ТРЕНД" if key == "тренд" else coll
    if "(6)" not in base:
        base = f"{base} (6)"
    return base


def lookup_oboi(artikul, full_name, catalog):
    key = artikul_as_key(artikul)
    if key in catalog:
        item = catalog[key]
        return item.get("code", ""), item.get("name", ""), True
    return "", oboi_heuristika_imeni(full_name), False


def _cell_str(v):
    if v is None:
        return ""
    return str(v).strip()


def eto_oboi_gotovyj(path):
    """Уже готовый УПД по обоям: есть коды УТ… в колонке товара."""
    name = os.path.basename(path).lower()
    if _yusuf_v_imeni(path) or eto_yusuf_gotovyj(path) or eto_yusuf_syroj(path):
        return False
    if "готов" in name:
        return True
    if path.lower().endswith("_обои_обработано.xlsx"):
        return True
    try:
        wb = load_workbook(path, data_only=True, read_only=True)
        ws = wb[wb.sheetnames[0]]
        ut_hits = 0
        oboi_hits = 0
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            if i > 120:
                break
            vals = [(_cell_str(v)) for v in row if v is not None]
            joined = " ".join(vals)
            if "Обои" in joined or "обои" in joined:
                oboi_hits += 1
            for v in vals[:3]:
                if _OBOI_CODE_RE.match(v):
                    ut_hits += 1
        wb.close()
        # Готовый эталон/выгрузка 1С: много УТ-кодов
        if ut_hits >= 3:
            return True
        return False
    except Exception:
        return False


def eto_oboi_syroj(path):
    """Сырой УПД/таблица по обоям (есть «Обои» в ячейках, нет УТ-кодов)."""
    if not path.lower().endswith(".xlsx"):
        return False
    if eto_oboi_gotovyj(path):
        return False
    if path.lower().endswith("_vitebsk_обработано.xlsx"):
        return False
    try:
        wb = load_workbook(path, data_only=True, read_only=True)
        ws = wb[wb.sheetnames[0]]
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            if i > 80:
                break
            for v in row:
                if v is not None and "обои" in str(v).lower():
                    wb.close()
                    return True
        wb.close()
    except Exception:
        return False
    # запасной признак — имя файла
    bn = os.path.basename(path).lower()
    return "упд" in bn or "ввп" in bn


def _find_oboi_header_row(ws):
    """Ищем строку с метками колонок 1а/1б/2/3… (предпочтительно) или шапку."""
    fallback = None
    for r in range(1, min(ws.max_row, 40) + 1):
        vals = {}
        for c in range(1, min(ws.max_column, 120) + 1):
            v = ws.cell(r, c).value
            if v is not None and str(v).strip() != "":
                vals[c] = str(v).strip()
        # классическая строка номеров колонок УПД — самый надёжный ориентир
        if any(v == "1а" for v in vals.values()) and any(v == "3" for v in vals.values()):
            return r, vals
        if fallback is None and any("Наименование товара" in v for v in vals.values()):
            fallback = (r, vals)
    if fallback is not None:
        return fallback
    return None, {}


def _map_oboi_columns(header_vals):
    """
    Возвращает dict логических полей → номер колонки.
    У сырого фрагмента и полного УПД раскладка разная.
    """
    inv = {v: c for c, v in header_vals.items()}

    def find_label(*variants):
        for var in variants:
            for c, v in header_vals.items():
                if v == var or var in v.replace("\n", " "):
                    return c
        return None

    # Вариант A: строка с 1а, 1б, 2, 2а, 3, …
    if "1а" in inv:
        cols = {
            "name": inv.get("1а"),
            "vid": inv.get("1б"),
            "ed_kod": inv.get("2"),
            "ed": inv.get("2а"),
            "qty": inv.get("3"),
            "price": inv.get("4"),
            "sum_wo": inv.get("5"),
            "excise": inv.get("6"),
            "rate": inv.get("7"),
            "nds": inv.get("8"),
            "sum_w": inv.get("9"),
            "country_code": inv.get("10"),
            "country": inv.get("10а"),
            "gtd": inv.get("11"),
            "code": None,
            "npp": None,
            "artikul": None,
        }
        return cols

    # Вариант B: полный УПД — ищем по подписям
    cols = {
        "code": find_label("Код товара"),
        "npp": find_label("№\nп/п", "№ п/п"),
        "name": find_label("Наименование товара"),
        "vid": find_label("Код\nвида", "Код вида"),
        "ed_kod": None,
        "ed": None,
        "qty": find_label("Коли-\nчество", "Количество"),
        "price": find_label("Цена"),
        "sum_wo": find_label("без налога"),
        "excise": find_label("акциза"),
        "rate": find_label("ставка"),
        "nds": find_label("Сумма налога"),
        "sum_w": find_label("с налогом"),
        "country_code": None,
        "country": None,
        "gtd": find_label("декларации"),
        "artikul": None,
    }
    return cols


def chitat_oboi_stroki(src_path):
    """Читает товарные строки сырого УПД/.xlsx по обоям."""
    wb = load_workbook(src_path, data_only=True)
    ws = wb[wb.sheetnames[0]]
    header_row, header_vals = _find_oboi_header_row(ws)
    if header_row is None:
        wb.close()
        raise ValueError(f"Не найден заголовок таблицы УПД в {os.path.basename(src_path)}")

    cols = _map_oboi_columns(header_vals)
    name_col = cols["name"]
    if not name_col:
        wb.close()
        raise ValueError("Не найдена колонка наименования в УПД")

    rows = []
    for r in range(header_row + 1, ws.max_row + 1):
        name = ws.cell(r, name_col).value
        if name is None or str(name).strip() == "":
            continue
        name_s = str(name).strip()
        if name_s.lower().startswith("всего"):
            break
        if "наименование" in name_s.lower():
            continue

        def get(key):
            c = cols.get(key)
            return ws.cell(r, c).value if c else None

        qty = get("qty")
        # строка данных должна иметь количество-число
        try:
            qtest = str(qty).replace(" ", "").replace(",", ".")
            float(qtest)
        except Exception:
            continue

        artikul, _ = extract_artikul_i_hvost(name_s)
        rows.append({
            "name_full": name_s,
            "artikul": artikul_as_key(artikul),
            "vid": get("vid"),
            "ed_kod": get("ed_kod"),
            "ed": get("ed"),
            "qty": qty,
            "price": get("price"),
            "sum_wo": get("sum_wo"),
            "excise": get("excise"),
            "rate": get("rate"),
            "nds": get("nds"),
            "sum_w": get("sum_w"),
            "country_code": get("country_code"),
            "country": get("country"),
            "gtd": get("gtd"),
        })
    wb.close()
    return rows


def transform_oboi_row(raw, npp, catalog):
    code, short_name, from_cat = lookup_oboi(raw["artikul"], raw["name_full"], catalog)
    return [
        code,
        npp,
        short_name,
        raw["artikul"],
        raw["vid"] if raw["vid"] is not None else "-",
        raw["ed_kod"],
        raw["ed"],
        raw["qty"],
        raw["price"],
        raw["sum_wo"],
        raw["excise"],
        raw["rate"],
        raw["nds"],
        raw["sum_w"],
        raw["country_code"] if raw["country_code"] is not None else "-",
        raw["country"] if raw["country"] is not None else "-",
        raw["gtd"] if raw["gtd"] is not None else "-",
    ], from_cat


def postroit_imya_rezultata_oboi(src_path):
    folder = os.path.dirname(src_path) or "."
    base = os.path.splitext(os.path.basename(src_path))[0]
    return unik_put(os.path.join(folder, f"{base}_ОБОИ_Обработано.xlsx"))


def convert_oboi(src_path, out_path=None, catalog=None):
    if catalog is None:
        catalog = load_oboi_catalog()
    raw_rows = chitat_oboi_stroki(src_path)
    processed = []
    known = 0
    unknown = 0
    for i, raw in enumerate(raw_rows, start=1):
        out_row, from_cat = transform_oboi_row(raw, i, catalog)
        processed.append(out_row)
        if from_cat:
            known += 1
        else:
            unknown += 1

    if out_path is None:
        out_path = postroit_imya_rezultata_oboi(src_path)

    log(f"Файл: {os.path.basename(src_path)}  [Обои / УПД]", "head")
    log(f"Обнаружено строк: {len(raw_rows)}")
    log(f"  - по справочнику oboi_catalog.json: {known}")
    if unknown:
        log(f"  - эвристика (артикул не в справочнике): {unknown}", "warn")
        log("    Проверьте эти позиции: короткое имя подобрано автоматически,", "warn")
        log("    код номенклатуры остался пустым.", "warn")
    log()

    wb_out = Workbook()
    ws = wb_out.active
    ws.title = "Обработанные данные"
    ws.append(OBOI_HEADER)
    for cell in ws[1]:
        cell.font = Font(name=FONT_NAME, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for out_row in processed:
        ws.append(out_row)
        for cell in ws[ws.max_row]:
            cell.font = Font(name=FONT_NAME)

    widths = [14, 10, 28, 14, 10, 10, 8, 10, 12, 14, 12, 10, 12, 14, 10, 12, 14]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = w
    ws.freeze_panes = "A2"
    wb_out.save(out_path)
    return len(raw_rows), out_path


def learn_oboi_catalog_from_gotovyj(path, folder=None):
    """
    Разбирает эталонный «готовый» УПД и дополняет oboi_catalog.json.
    Ожидает колонки: код УТ…, короткое имя, артикул.
    """
    folder = folder or script_dir()
    catalog = load_oboi_catalog(folder)
    wb = load_workbook(path, data_only=True)
    ws = wb[wb.sheetnames[0]]

    added = 0
    updated = 0
    # Ищем строки, где в первых колонках код УТ и дальше имя+артикул
    for r in range(1, ws.max_row + 1):
        code = None
        name = None
        artikul = None
        # Сканируем строку
        cells = []
        for c in range(1, min(ws.max_column, 120) + 1):
            v = ws.cell(r, c).value
            if v is not None and str(v).strip() != "":
                cells.append((c, v))
        if not cells:
            continue
        # код
        for c, v in cells:
            if isinstance(v, str) and _OBOI_CODE_RE.match(v.strip()):
                code = v.strip()
                break
        if not code:
            continue
        # артикул: число или токен вида 1094-01ОАВ рядом с именем
        for c, v in cells:
            key = artikul_as_key(v)
            if re.fullmatch(r"\d{4,}", key) or re.fullmatch(r"\d{3,}-\d{2}[A-Za-zА-Яа-яЁё]+", key):
                # не путать с кодом страны 112 и т.п. — артикул обычно >= 4 цифр или с дефисом
                if key == "112":
                    continue
                artikul = key
                # имя — предыдущая непустая ячейка до артикула, не код и не номер строки
                prev_candidates = [vv for cc, vv in cells if cc < c]
                for vv in reversed(prev_candidates):
                    s = str(vv).strip()
                    if s == code:
                        continue
                    if re.fullmatch(r"\d+", s) and int(s) < 1000:
                        continue  # № п/п
                    if _OBOI_CODE_RE.match(s):
                        continue
                    name = s
                    break
                break
        if not artikul or not name:
            continue
        prev = catalog.get(artikul)
        catalog[artikul] = {"code": code, "name": name}
        if prev is None:
            added += 1
        elif prev != catalog[artikul]:
            updated += 1

    wb.close()
    save_oboi_catalog(catalog, folder)
    log(f"Справочник обоев: +{added} новых, {updated} обновлено, всего {len(catalog)}", "ok")
    log(f"Файл: {os.path.join(folder, OBOI_CATALOG_FILENAME)}")
    return catalog


# ===========================================================================
# Журнал / поиск файлов / main
# ===========================================================================
def load_state(folder):
    path = os.path.join(folder, STATE_FILENAME)
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def save_state(folder, state):
    path = os.path.join(folder, STATE_FILENAME)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)


def opredelit_tip_faila(path):
    """Возвращает 'vitebsk' | 'oboi' | 'yusuf' | 'skip'."""
    low = path.lower()
    base = os.path.basename(low)
    if base.startswith("~$"):
        return "skip"
    if "_обработано" in base:
        return "skip"
    if low.endswith(".xls") and not low.endswith(".xlsx"):
        if eto_gotovyj_vitebsk(path):
            return "skip"
        return "vitebsk"
    if low.endswith(".xlsx"):
        if eto_yusuf_gotovyj(path):
            return "skip"
        if eto_yusuf_syroj(path):
            return "yusuf"
        if eto_oboi_gotovyj(path):
            return "skip"
        if eto_oboi_syroj(path):
            return "oboi"
        return "skip"
    return "skip"


def find_source_files(folder=".", state=None):
    """Новые накладные в папке. Возвращает (список (путь, тип), уже обработанные, нераспознанные)."""
    candidates = []
    already = []
    unknown = []
    for name in sorted(os.listdir(folder)):
        low = name.lower()
        if not (low.endswith(".xls") or low.endswith(".xlsx")):
            continue
        if name.startswith("~$") or name.startswith("."):
            continue
        path = os.path.join(folder, name)
        if not os.path.isfile(path):
            continue
        tip = opredelit_tip_faila(path)
        if tip == "skip":
            if "_обработано" not in low and "готов" not in low:
                unknown.append(name)
            continue
        if state is not None and name in state:
            already.append((name, state[name]))
            continue
        candidates.append((path, tip))
    return candidates, already, unknown


def convert_any(src_path, out_path=None):
    tip = opredelit_tip_faila(src_path)
    # Явно указанный файл: если «skip» из-за «готовый» — всё равно пробуем угадать
    if tip == "skip":
        low = src_path.lower()
        if low.endswith(".xls") and not low.endswith(".xlsx"):
            tip = "vitebsk"
        elif low.endswith(".xlsx"):
            if eto_yusuf_syroj(src_path) or eto_packing_list(src_path) or _yusuf_v_imeni(src_path):
                tip = "yusuf"
            else:
                tip = "oboi"
        else:
            raise ValueError(f"Непонятный тип файла: {src_path}")
    if tip == "vitebsk":
        return convert_vitebsk(src_path, out_path) + ("vitebsk",)
    if tip == "oboi":
        return convert_oboi(src_path, out_path) + ("oboi",)
    if tip == "yusuf":
        return convert_yusuf(src_path, out_path) + ("yusuf",)
    raise ValueError(f"Не удалось определить тип: {src_path}")


def zapisat_v_zhurnal(folder_state, state, src_name, out_path, rows, tip):
    state[src_name] = {
        "processed_at": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "out_file": os.path.basename(out_path),
        "rows": rows,
        "type": tip,
    }
    save_state(folder_state, state)


# ===========================================================================
# Точки входа для GUI и командной строки
# ===========================================================================
def process_folder(folder=None, force=False):
    """
    Обрабатывает все новые накладные в папке. Пишет ход работы через log()
    и возвращает сводку для показа пользователю.
    """
    folder = os.path.abspath(folder or app_dir())
    started = datetime.datetime.now()
    itog = {
        "folder": folder,
        "done": [],
        "skipped": [],
        "errors": [],
        "rows": 0,
        "seconds": 0.0,
    }

    if not os.path.isdir(folder):
        log(f"Папка не найдена: {folder}", "err")
        itog["errors"].append(("", f"Папка не найдена: {folder}"))
        return itog

    log(f"Папка: {folder}")
    catalog = load_oboi_catalog()
    if catalog:
        log(f"Справочник обоев: {len(catalog)} артикулов")
    else:
        log(f"Справочник обоев не найден ({OBOI_CATALOG_FILENAME}) — короткие имена", "warn")
        log("будут подобраны эвристикой, коды номенклатуры останутся пустыми.", "warn")
    if force:
        log("Режим повторной обработки: журнал обработанных файлов игнорируется.", "warn")
    log()

    state = load_state(folder)
    sources, already, unknown = find_source_files(folder, state=None if force else state)

    for name, prev in already:
        itog["skipped"].append((name, prev))
        log(f"Пропускаю (обработан {prev.get('processed_at', '?')} -> "
            f"{prev.get('out_file', '?')}): {name}")
    if already:
        log()

    if unknown:
        log("Не распознаны (пропущены):", "warn")
        for name in unknown:
            log(f"  • {name}", "warn")
        log("Нужен витебский .xls, сырой УПД обоев или packing list Юсуф.", "warn")
        log()

    if not sources:
        log("Новых накладных (.xls / .xlsx) в этой папке нет.", "warn")
        log("Витебские ковры: положите .xls рядом с приложением.")
        log("Обои: положите сырой УПД .xlsx (со словом «Обои» в наименованиях).")
        log("Юсуф: положите packing list .xlsx (ковры, Турция).")
        if already:
            log("Чтобы обработать файлы заново, включите «Обрабатывать повторно».")
        itog["seconds"] = (datetime.datetime.now() - started).total_seconds()
        return itog

    log(f"К обработке: {len(sources)} файл(ов)", "head")
    log()

    for i, (src_path, tip) in enumerate(sources, start=1):
        src_name = os.path.basename(src_path)
        log(f"[{i}/{len(sources)}] {src_name}", "head")
        try:
            if tip == "vitebsk":
                rows, out_path = convert_vitebsk(src_path)
            elif tip == "yusuf":
                rows, out_path = convert_yusuf(src_path)
            else:
                rows, out_path = convert_oboi(src_path, catalog=catalog)
        except Exception as exc:
            itog["errors"].append((src_name, str(exc)))
            log(f"ОШИБКА: {exc}", "err")
            log()
            continue

        itog["done"].append({
            "src": src_name,
            "out": out_path,
            "rows": rows,
            "type": tip,
        })
        itog["rows"] += rows
        log(f"Готово: {rows} строк -> {os.path.basename(out_path)}", "ok")
        zapisat_v_zhurnal(folder, state, src_name, out_path, rows, tip)
        log("-" * 60)

    itog["seconds"] = (datetime.datetime.now() - started).total_seconds()
    return itog


def process_one_file(src_path, out_path=None, force=False, folder_state=None):
    """Обрабатывает один явно указанный файл. Возвращает такую же сводку."""
    folder_state = os.path.abspath(folder_state or os.path.dirname(os.path.abspath(src_path)))
    started = datetime.datetime.now()
    itog = {
        "folder": folder_state,
        "done": [],
        "skipped": [],
        "errors": [],
        "rows": 0,
        "seconds": 0.0,
    }
    src_name = os.path.basename(src_path)
    state = load_state(folder_state)

    if not os.path.exists(src_path):
        log(f"Файл не найден: {src_path}", "err")
        itog["errors"].append((src_name, "файл не найден"))
        return itog

    if not force and src_name in state:
        prev = state[src_name]
        itog["skipped"].append((src_name, prev))
        log(f"Этот файл уже обработан {prev.get('processed_at', '?')} -> "
            f"{prev.get('out_file', '?')}", "warn")
        log("Для повторной обработки добавьте --force.")
        itog["seconds"] = (datetime.datetime.now() - started).total_seconds()
        return itog

    try:
        rows, out_path, tip = convert_any(src_path, out_path)
    except Exception as exc:
        itog["errors"].append((src_name, str(exc)))
        log(f"ОШИБКА: {exc}", "err")
        itog["seconds"] = (datetime.datetime.now() - started).total_seconds()
        return itog

    itog["done"].append({"src": src_name, "out": out_path, "rows": rows, "type": tip})
    itog["rows"] = rows
    log(f"Готово: {rows} строк перенесено -> {out_path}", "ok")
    zapisat_v_zhurnal(folder_state, state, src_name, out_path, rows, tip)
    itog["seconds"] = (datetime.datetime.now() - started).total_seconds()
    return itog


def opisanie_itoga(itog):
    """Короткая строка о результате — для статуса в GUI и финальной строки в консоли."""
    if itog["errors"]:
        return f"Завершено с ошибками: {len(itog['errors'])} из {len(itog['errors']) + len(itog['done'])}"
    if itog["done"]:
        files = len(itog["done"])
        return f"Готово: обработано файлов — {files}, строк — {itog['rows']}"
    if itog["skipped"]:
        return "Новых файлов нет: все накладные уже обработаны"
    return "Новых накладных в папке не найдено"


def _arg_value(raw_args, flag):
    if flag in raw_args:
        idx = raw_args.index(flag)
        if idx + 1 < len(raw_args):
            return raw_args[idx + 1]
    prefix = flag + "="
    for a in raw_args:
        if a.startswith(prefix):
            return a[len(prefix):]
    return None


def main(argv=None):
    raw_args = list(sys.argv[1:] if argv is None else argv)
    force = "--force" in raw_args
    flags = ("--force", "--learn-oboi", "--learn-inbox", "--setup-inbox", "--watch",
             "--folder")
    args = []
    skip_next = False
    for i, a in enumerate(raw_args):
        if skip_next:
            skip_next = False
            continue
        if a in flags:
            if a in ("--learn-oboi", "--folder") and i + 1 < len(raw_args):
                skip_next = True
            continue
        if a.startswith("--folder=") or a.startswith("--learn-oboi="):
            continue
        args.append(a)

    if "--setup-inbox" in raw_args or "--learn-inbox" in raw_args or "--watch" in raw_args:
        import learn_novye
        inbox = _arg_value(raw_args, "--folder") or learn_novye.default_novye_dir()
        if "--setup-inbox" in raw_args:
            path = learn_novye.sozdat_papku_novye(inbox)
            log("Папка готова: %s" % path, "ok")
            if "--learn-inbox" not in raw_args and "--watch" not in raw_args:
                return 0
        if "--watch" in raw_args:
            return learn_novye.watch_inbox(inbox) or 0
        itog = learn_novye.process_inbox(inbox, force=force)
        log()
        log(opisanie_itoga(itog), "err" if itog["errors"] else "ok")
        return 1 if itog["errors"] else 0

    if "--learn-oboi" in raw_args:
        try:
            idx = raw_args.index("--learn-oboi")
            learn_path = raw_args[idx + 1]
        except Exception:
            log("Укажите файл: python convert_nakladnaya.py --learn-oboi \"… готовый .xlsx\"", "err")
            return 1
        learn_oboi_catalog_from_gotovyj(learn_path, folder=app_dir())
        return 0

    # Накладные лежат в текущей папке (bat/exe запускаются из рабочего каталога)
    folder_state = os.path.abspath(".")

    if args:
        itog = process_one_file(args[0], args[1] if len(args) >= 2 else None,
                                force=force, folder_state=folder_state)
    else:
        itog = process_folder(folder_state, force=force)

    log()
    log(opisanie_itoga(itog), "err" if itog["errors"] else "ok")
    return 1 if itog["errors"] else 0


if __name__ == "__main__":
    sys.exit(main())
