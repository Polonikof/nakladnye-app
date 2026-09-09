# -*- coding: utf-8 -*-
"""
VtebskieKovri — обработка накладных от Витебских ковров в формат
номенклатуры вашей программы.

Термин "коллекция" (не "бренд") — название линейки товара, зашитое
в столбец "Рисунок/Колорит". Коллекций больше, чем перечислено ниже;
известны только несколько особых, для остальных используется общий
алгоритм — прямой перевод вашего VBA-макроса.

Определение коллекции по строке "Рисунок/Колорит":

  ПАЛИТРА
     Формат "pАртикул/Цветp/Серия", например "p2638/a4r/124" — ведущая
     буква 'p' перед цифрой. Внутри "мусорные" буквы r/, p/, // —
     убираются. Наименование = очищенная строка целиком, например
     "2638/a4/124". Ковролин.

  ШЕГГИ
     Формат "sh/x/Серия", например "sh/r/34", "sh/p/85" — начинается
     с "sh/". Та же чистка мусорных букв, что и у Палитры, но без
     удаления "sh" (это не баг, а часть кода коллекции). Наименование,
     например "sh/34". Ковролин.

  Именная коллекция (РОКСОЛАНА, КОНСОНАНС и любая другая новая)
     Формат "Артикул/Цвет/КОЛЛЕКЦИЯ", например "51281/a6/РОКСОЛАНА".
     Категория (ковёр/ковролин) определяется по столбцу "Продукция":
     если есть слово "ковёр"/"ковер" -> ковёр, иначе -> ковролин
     (например "Дорожка ...").
       - Ковёр: Наименование = КОЛЛЕКЦИЯ + размеры + форма.
       - Ковролин: Наименование = только КОЛЛЕКЦИЯ; ПолноеНаименование
         = КОЛЛЕКЦИЯ + первое слово из "Продукция"; Характеристика =
         Артикул+Цвет+Ширина+" м".

  Неопознанная коллекция — всё, что не подошло ни под один из
  форматов выше. Обрабатывается ОБЩИМ алгоритмом — прямым переводом
  исходного VBA-макроса поставщика (порог "Ковролин, если длина > 5 м",
  распознавание Овала по 'o/' или '/ЭФ', разбор "//" в наименовании
  и характеристике). Название коллекции в имя результата не попадает.

Готовый файл называется:
    Vitebsk_<дата обработки>_<коллекция1>_<коллекция2>...xlsx
Если коллекция не определена ни у одной строки — часть с коллекциями
опускается: Vitebsk_<дата>.xlsx

Использование:
    python VtebskieKovri.py                      -> обработать все .xls в папке
    python VtebskieKovri.py <файл.xls>            -> обработать конкретный файл
    python VtebskieKovri.py <файл.xls> <рез.xlsx> -> указать имя результата явно
"""

import sys
import os
import datetime
import json
import xlrd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

# Файл-журнал уже обработанных накладных (лежит рядом со скриптом).
# Ключ - ИМЯ исходного файла (os.path.basename). Если файл переименуют,
# он будет считаться новым и обработается снова - это осознанный выбор
# (проверка сделана простой и предсказуемой, по имени файла).
STATE_FILENAME = ".vitebsk_processed.json"

HEADER = [
    "НомерСтроки", "Тип", "Форма", "Наименование", "ПолноеНаименвоание",
    "ЕдиницаБазовая", "ЕдиницаХранения", "ЕдиницаХраненияКоэфф",
    "Характеристика", "Серия", "Ширина", "Длина", "Штрихкод",
    "Количество", "Сумма",
]

FONT_NAME = "Arial"

# Порог длины (м) для общего VBA-алгоритма (неопознанные коллекции):
# длиннее — рулонный товар (Ковролин), короче — штучный (Ковер).
VBA_DLINA_KOVROLIN_POROG = 5


def norm_num(v):
    """Число без лишнего .0, с запятой вместо точки — как в Excel RU."""
    if isinstance(v, float) and v.is_integer():
        s = str(int(v))
    else:
        s = str(v)
    return s.replace('.', ',')


def to_int_if_whole(v):
    if isinstance(v, float) and v.is_integer():
        return int(v)
    return v


def pervoe_slovo(text):
    s = str(text).strip()
    if not s:
        return ''
    return s.split()[0]


def opredelit_kategoriya(produkciya):
    """Категория товара 'ковер'/'ковролин' по столбцу 'Продукция'
    (используется для именных коллекций типа РОКСОЛАНА/КОНСОНАНС)."""
    s = str(produkciya).lower().replace('ё', 'е')
    if 'ковер' in s:
        return "ковер"
    return "ковролин"


import re

_CYRILLIC_RE = re.compile(r'[а-яёА-ЯЁ]')


def opredelit_kollekciyu(risunok):
    """Определяет коллекцию по строке 'Рисунок/Колорит'.
    Возвращает: 'Палитра', 'Шегги', название именной коллекции
    (например 'РОКСОЛАНА') или None, если формат не распознан
    (тогда применяется общий VBA-алгоритм).

    ВАЖНО: именная коллекция отличается от простого 2-3-буквенного
    цветового/паттерн-кода (например 'es', 'ct', 'ov', 'zf', 'cas' —
    встречаются в накладных с "Дорожками") тем, что записывается
    кириллицей заглавными буквами (РОКСОЛАНА, КОНСОНАНС). Код цвета/
    паттерна пишется латиницей строчными буквами и коллекцией не
    является — такие строки должны обрабатываться общим алгоритмом
    (как Палитра/Шегги), а не как именная коллекция. Раньше здесь
    проверялось только "3-й элемент не цифра", из-за чего 'es'/'ct'/
    'ov'/'zf'/'cas' ошибочно принимались за названия коллекций —
    это и было причиной некорректной обработки таких накладных.
    """
    s = str(risunok).strip()
    if len(s) > 1 and s[0] == 'p' and s[1].isdigit():
        return "Палитра"
    if s.lower().startswith('sh/'):
        return "Шегги"
    parts = [p.strip() for p in s.split('/')]
    if len(parts) >= 3 and parts[2] and _CYRILLIC_RE.search(parts[2]):
        return parts[2].upper()
    return None


# ---------------------------------------------------------------------
# Именная коллекция (РОКСОЛАНА, КОНСОНАНС, ...), категория "ковер"
# ---------------------------------------------------------------------
def transform_kover_imennoy(row):
    (num, produkciya, izdelie, risunok, shirina, dlina, sort, razbrak,
     kolvo, ed, ploshad, cena, summa, shtrih) = row[:14]

    parts = [p.strip() for p in str(risunok).split('/')]
    artikul = parts[0] if len(parts) > 0 else ''
    tsvet = parts[1] if len(parts) > 1 else ''
    kollekciya = parts[2] if len(parts) > 2 else ''

    tip = "Ковер"

    if 'o/' in str(risunok):
        forma = "Овал"
    elif '/ЭФ' in str(risunok):
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


# ---------------------------------------------------------------------
# Именная коллекция (РОКСОЛАНА, КОНСОНАНС, ...), категория "ковролин"
# (например "Дорожка")
# ---------------------------------------------------------------------
def transform_kovrolin_imennoy(row, kollekciya):
    (num, produkciya, izdelie, risunok, shirina, dlina, sort, razbrak,
     kolvo, ed, ploshad, cena, summa, shtrih) = row[:14]

    parts = [p.strip() for p in str(risunok).split('/')]
    artikul = parts[0] if len(parts) > 0 else ''
    tsvet = parts[1] if len(parts) > 1 else ''

    tip = "Ковролин"
    forma = ""

    naimenovanie = kollekciya
    polnoe = f"{kollekciya} {pervoe_slovo(produkciya)}".strip()

    w_str = norm_num(shirina)
    harakteristika = f"{artikul} {tsvet} {w_str} м".strip()

    seria = to_int_if_whole(izdelie)
    kolichestvo = ploshad

    return [
        to_int_if_whole(num), tip, forma, naimenovanie, polnoe,
        "м2", "м2", 1,
        harakteristika, seria,
        shirina, dlina, shtrih, kolichestvo, summa,
    ]


# ---------------------------------------------------------------------
# Ковролин ПАЛИТРА и ШЕГГИ — одна и та же схема очистки строки
# ---------------------------------------------------------------------
def ochistit_stroku_kovrolin(risunok):
    """Убирает служебные буквы r/, p/, // и (только для Палитры)
    ведущую 'p' перед цифрой артикула."""
    clean = str(risunok).strip()
    clean = clean.replace('/Шегги ', '').replace('Шегги ', '')
    clean = clean.replace('r/', '/').replace('p/', '/').replace('//', '/')
    # Ведущая буква "p" перед цифрой — брак исходных данных Палитры,
    # убираем её. У Шегги строка начинается с "sh", это не затрагивает.
    if len(clean) > 1 and clean[0] == 'p' and clean[1].isdigit():
        clean = clean[1:]
    return clean


def transform_kovrolin_ochistka(row):
    """Ковролин ПАЛИТРА или ШЕГГИ — схема как в VBA-макросе поставщика."""
    (num, produkciya, izdelie, risunok, shirina, dlina, sort, razbrak,
     kolvo, ed, ploshad, cena, summa, shtrih) = row[:14]

    tip = "Ковролин"
    forma = ""

    naimenovanie = ochistit_stroku_kovrolin(risunok)
    polnoe = f"{tip} {naimenovanie}"

    w_str = norm_num(shirina)
    harakteristika = f"{w_str} м"

    seria = to_int_if_whole(izdelie)
    kolichestvo = ploshad  # Количество для ковролина = Площадь, м2

    return [
        to_int_if_whole(num), tip, forma, naimenovanie, polnoe,
        "м2", "м2", 1,
        harakteristika, seria,
        shirina, dlina, shtrih, kolichestvo, summa,
    ]


# ---------------------------------------------------------------------
# Неопознанная коллекция — общий алгоритм, прямой перевод VBA-макроса
# ---------------------------------------------------------------------
def transform_vba_obshiy(row):
    (num, produkciya, izdelie, risunok, shirina, dlina, sort, razbrak,
     kolvo, ed, ploshad, cena, summa, shtrih) = row[:14]

    pattern = str(risunok).strip()
    tip = "Ковролин" if dlina > VBA_DLINA_KOVROLIN_POROG else "Ковер"

    if tip == "Ковер":
        if 'o/' in pattern:
            forma = "Овал"
        elif '/ЭФ' in pattern:
            forma = "Овал"
        elif shirina == dlina:
            forma = "Круг"
        else:
            forma = "Прямоугольник"
    else:
        forma = ""

    # --- Наименование (перевод VBA) ---
    if tip == "Ковер":
        temp = pattern.replace('/Шегги ', '')
        if '//' in temp:
            parts = temp.split('//')
            if len(parts) >= 2:
                name_parts = parts[1].strip().split(' ')
                temp = name_parts[0] if name_parts and name_parts[0] else parts[1].strip()
        w_str = norm_num(shirina)
        l_str = norm_num(dlina)
        temp = f"{temp} {w_str}*{l_str} {forma}"
    else:
        temp = pattern.replace('Шегги ', '')

    temp = temp.replace('r/', '/').replace('p/', '/').replace('//', '/')
    naimenovanie = temp
    polnoe = f"{tip} {naimenovanie}"

    # --- Характеристика (перевод VBA) ---
    if tip == "Ковролин":
        w_str = norm_num(shirina)
        harakteristika = f"{w_str} м"
    else:
        harakteristika = ""
        if '//' in pattern:
            parts = pattern.split('//')
            if len(parts) >= 2:
                before = parts[0].strip()
                after = parts[1].strip()
                name_parts = after.split(' ')
                if len(name_parts) >= 2:
                    ostatok = after.replace(name_parts[0], '', 1).strip()
                    harakteristika = f"{before} {ostatok}".strip()
                else:
                    harakteristika = before

    # --- Прочие поля (перевод VBA) ---
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
    """Возвращает (готовая_строка, коллекция_или_None)."""
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

    # Неопознанный формат — общий VBA-алгоритм, коллекция не определена
    return transform_vba_obshiy(row), None


def chitat_syrye_stroki(src_path):
    """Читает исходник и возвращает список сырых строк с данными
    (пропуская пустые/итоговые строки без номера или без нормальных
    ширины/длины)."""
    wb_src = xlrd.open_workbook(src_path)
    sh = wb_src.sheet_by_index(0)

    raw_rows = []
    for r in range(1, sh.nrows):
        raw = [sh.cell_value(r, c) for c in range(sh.ncols)]
        if raw[0] == '' or raw[0] is None:
            continue
        if not isinstance(raw[4], (int, float)) or not isinstance(raw[5], (int, float)):
            continue
        raw_rows.append(raw)
    return raw_rows


def sanitize_dlya_imeni(text):
    """Делает строку безопасной для использования в имени файла Windows."""
    bad = '<>:"/\\|?*'
    s = str(text)
    for ch in bad:
        s = s.replace(ch, '_')
    return s.replace(' ', '_')


def postroit_imya_rezultata(src_path):
    """<имя_исходника>_VITEBSK_Обработано.xlsx
    Например: 'Исходник_ФактураЕАН13_1252227.xls' ->
    'Исходник_ФактураЕАН13_1252227_VITEBSK_Обработано.xlsx'."""
    folder = os.path.dirname(src_path) or "."
    base = os.path.splitext(os.path.basename(src_path))[0]
    name = f"{base}_VITEBSK_Обработано.xlsx"

    path = os.path.join(folder, name)
    # избегаем перезаписи, если такой файл уже есть (например, второй
    # запуск с --force по тому же исходнику)
    if os.path.exists(path):
        base_path, ext = os.path.splitext(path)
        i = 2
        while os.path.exists(f"{base_path} ({i}){ext}"):
            i += 1
        path = f"{base_path} ({i}){ext}"
    return path


def eto_gotovyj_fail(path):
    """True, если файл — уже готовый/эталонный результат (в нём есть
    лист 'Обработанные данные'), а не сырой исходник от поставщика.
    Так надёжнее, чем ориентироваться на слова вроде 'копия' в имени —
    исходник может называться как угодно."""
    try:
        wb = xlrd.open_workbook(path)
        return "Обработанные данные" in wb.sheet_names()
    except Exception:
        return True


def convert(src_path, out_path=None):
    raw_rows = chitat_syrye_stroki(src_path)

    kollekcii = []          # уникальные коллекции, в порядке появления
    counts = {}              # {метка: кол-во строк} для сообщения
    processed = []            # (готовая_строка) по каждой строке

    for raw in raw_rows:
        out_row, kollekciya = transform_row(raw)
        processed.append(out_row)

        if kollekciya is not None and kollekciya not in kollekcii:
            kollekcii.append(kollekciya)

        label = kollekciya if kollekciya is not None else "не определена (общий VBA-алгоритм)"
        counts[label] = counts.get(label, 0) + 1

    if out_path is None:
        out_path = postroit_imya_rezultata(src_path)

    print(f"Файл: {os.path.basename(src_path)}")
    print("Обнаружено:")
    for label, cnt in sorted(counts.items(), key=lambda x: -x[1]):
        print(f"  - {label}: {cnt} стр.")
    if len(kollekcii) > 1:
        print(f"В файле несколько коллекций сразу — все они попадут в один")
        print(f"итоговый файл: {', '.join(kollekcii)}")
    print()

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
        # Штрихкод (13-я колонка) - явный формат "0" (целое число без
        # разделителей и без экспоненциальной записи). Без этого Excel
        # при узкой колонке показывает штрихкод как "4.81034E+12" -
        # именно так, как в эталонном файле ФактураЕАН13.
        ws.cell(row=ws.max_row, column=13).number_format = "0"

    widths = [10, 10, 14, 26, 32, 12, 14, 16, 14, 10, 9, 9, 16, 11, 10]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = w
    ws.freeze_panes = "A2"

    wb_out.save(out_path)
    return len(raw_rows), out_path


def load_state(folder):
    path = os.path.join(folder, STATE_FILENAME)
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            # журнал повреждён/нечитаем - лучше начать заново, чем упасть
            return {}
    return {}


def save_state(folder, state):
    path = os.path.join(folder, STATE_FILENAME)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)


def find_source_files(folder=".", state=None):
    """Находит файлы-исходники в папке: любой .xls, который
    (а) не является уже готовым/эталонным файлом (см. eto_gotovyj_fail) и
    (б) ещё не встречался в журнале обработанных файлов (state) -
        проверка по ИМЕНИ файла."""
    candidates = []
    for name in os.listdir(folder):
        if not name.lower().endswith(".xls"):
            continue
        path = os.path.join(folder, name)
        if eto_gotovyj_fail(path):
            continue
        if state is not None and name in state:
            prev = state[name]
            print(f"Пропускаю (уже обработан {prev.get('processed_at', '?')}, "
                  f"результат {prev.get('out_file', '?')}): {name}")
            continue
        candidates.append(path)
    return candidates


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if a != "--force"]
    force = "--force" in sys.argv

    folder = "."
    state = load_state(folder)

    if len(args) >= 1:
        src_path = args[0]
        out_path = args[1] if len(args) >= 2 else None
        src_name = os.path.basename(src_path)

        if not force and src_name in state:
            prev = state[src_name]
            print(f"Этот файл уже был обработан {prev.get('processed_at', '?')} "
                  f"-> {prev.get('out_file', '?')}")
            print("Чтобы обработать повторно, добавьте флаг --force:")
            print(f"  python VtebskieKovri.py \"{src_path}\" --force")
            sys.exit(0)

        n, out_path = convert(src_path, out_path)
        print(f"Готово: {n} строк перенесено -> {out_path}")

        state[src_name] = {
            "processed_at": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "out_file": os.path.basename(out_path),
            "rows": n,
        }
        save_state(folder, state)
    else:
        sources = find_source_files(folder, state=None if force else state)
        if not sources:
            print("В этой папке не найдено новых накладных (.xls).")
            print("Либо все .xls уже были обработаны ранее (см. журнал),")
            print("либо положите файл рядом со скриптом и запустите снова.")
            print("Указать файл явно: python VtebskieKovri.py \"файл.xls\"")
            print("Обработать повторно: python VtebskieKovri.py --force")
            sys.exit(0)

        for src_path in sources:
            n, out_path = convert(src_path)
            print(f"Готово: {n} строк -> {os.path.basename(out_path)}")

            src_name = os.path.basename(src_path)
            state[src_name] = {
                "processed_at": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "out_file": os.path.basename(out_path),
                "rows": n,
            }
            save_state(folder, state)
            print("-" * 40)
