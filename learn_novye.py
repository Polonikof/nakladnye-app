# -*- coding: utf-8 -*-
"""
Агент обучения по папке «НовыеНакладные».

Вы кладёте пару файлов:
  • исходник (как пришло от поставщика)
  • готовый (как обработали руками)

Агент находит пару, разбирает её:
  • Обои — дополняет oboi_catalog.json
  • Юсуф — запоминает формат packing list (колонки наим / харка)
  • Витебск — архивирует эталон

и готовит данные для нового релиза Накладные.exe.

    python learn_novye.py --setup
    python learn_novye.py --once
    python learn_novye.py --watch
    python convert_nakladnaya.py --learn-inbox
"""

from __future__ import print_function

import argparse
import datetime
import hashlib
import json
import os
import re
import shutil
import sys
import time

import convert_nakladnaya as core
from convert_nakladnaya import log

STATE_NAME = ".learned.json"
ARCHIVE_DIR = "_архив"
INSTRUCTION_NAME = "КАК_РАБОТАТЬ.txt"

SOURCE_MARKERS = ("source", "исход", "сырой", "вход")
READY_MARKERS = ("готов", "result", "эталон", "образец")
EXCEL_EXT = (".xls", ".xlsx")

INSTRUCTION = """КАК РАБОТАТЬ

Папка: C:\\Накладные\\НовыеНакладные

1. Положите в эту папку ДВА файла одной накладной:
   - как пришло от поставщика (в имени напишите слово исходник)
   - как сделали руками        (в имени напишите слово готовый)

   Примеры:
   УПД_май_исходник.xlsx
   УПД_май_готовый.xlsx

   Юсуф исходник №1643.xlsx
   Готовый Юсуф №1643.xlsx

   Если исходник тот же, а готовый файл другой (или вы поправили
   готовый и сохранили), правила этого типа накладных перезапишутся.

2. Откройте Накладные.exe (лежит в C:\\Накладные).

3. Нажмите кнопку «Обучить».

4. Готово. Появится один файл Накладные_1.6.1.exe
   в C:\\Накладные. Его можно сразу отправить коллеге:
   все правила уже внутри, больше ничего не нужно.
   Закройте старое окно и откройте этот файл.
"""


def default_work_dir():
    if sys.platform.startswith("win") and os.path.isdir(r"C:\Накладные"):
        return r"C:\Накладные"
    return core.app_dir()


def default_novye_dir():
    env = os.environ.get("NAKLADNYE_NOVYE")
    if env:
        return os.path.abspath(env)
    beside = os.path.join(core.app_dir(), "НовыеНакладные")
    if os.path.isdir(beside):
        return beside
    preferred = r"C:\Накладные\НовыеНакладные"
    if sys.platform.startswith("win"):
        return preferred
    return os.path.join(os.path.expanduser("~"), "НовыеНакладные")


def sozdat_papku_novye(folder=None):
    folder = os.path.abspath(folder or default_novye_dir())
    try:
        os.makedirs(folder, exist_ok=True)
    except OSError:
        fallback = os.path.join(os.path.expanduser("~"), "НовыеНакладные")
        log("Не удалось создать %s — использую %s" % (folder, fallback), "warn")
        folder = fallback
        os.makedirs(folder, exist_ok=True)
    os.makedirs(os.path.join(folder, ARCHIVE_DIR), exist_ok=True)
    instr = os.path.join(folder, INSTRUCTION_NAME)
    if not os.path.exists(instr):
        with open(instr, "w", encoding="utf-8") as f:
            f.write(INSTRUCTION)
    return folder


def _role_by_name(name):
    n = name.lower()
    if any(m in n for m in READY_MARKERS):
        return "gotovyj"
    if any(m in n for m in SOURCE_MARKERS):
        return "source"
    return "unknown"


def _klass_puti(path):
    """source_oboi | gotovyj_oboi | source_yusuf | gotovyj_yusuf | source_vitebsk | skip"""
    base = os.path.basename(path)
    if base.startswith("~$") or base.startswith("."):
        return "skip"
    low = path.lower()
    if ARCHIVE_DIR.lower() in low.replace("\\", "/").split("/"):
        return "skip"
    if not low.endswith(EXCEL_EXT):
        return "skip"
    by_name = _role_by_name(base)
    if low.endswith(".xls") and not low.endswith(".xlsx"):
        if core.eto_gotovyj_vitebsk(path):
            return "gotovyj_vitebsk"
        return "source_vitebsk" if by_name != "gotovyj" else "gotovyj_vitebsk"
    if core.eto_yusuf_gotovyj(path) or (
            by_name == "gotovyj" and (core._yusuf_v_imeni(path) or core.eto_packing_list(path))):
        return "gotovyj_yusuf"
    if core.eto_yusuf_syroj(path) or (
            by_name == "source" and (core._yusuf_v_imeni(path) or core.eto_packing_list(path))):
        return "source_yusuf"
    if core.eto_oboi_gotovyj(path) or by_name == "gotovyj":
        return "gotovyj_oboi"
    if core.eto_oboi_syroj(path) or by_name == "source":
        return "source_oboi"
    return "skip"


def _stem(name):
    n = os.path.splitext(name)[0].lower()
    for m in SOURCE_MARKERS + READY_MARKERS:
        n = n.replace(m, " ")
    n = re.sub(r"[^a-zа-яё0-9]+", " ", n)
    return re.sub(r"\s+", " ", n).strip()


def _stem_score(a, b):
    ta = set(_stem(a).split())
    tb = set(_stem(b).split())
    if not ta or not tb:
        return 0.0
    return len(ta & tb) / float(len(ta | tb))


def _file_sig(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def _mtime(path):
    try:
        return os.path.getmtime(path)
    except OSError:
        return 0.0


def _excel_files(folder):
    out = []
    for root, dirs, files in os.walk(folder):
        dirs[:] = [d for d in dirs if d != ARCHIVE_DIR and not d.startswith(".")]
        for name in files:
            path = os.path.join(root, name)
            klass = _klass_puti(path)
            if klass != "skip":
                out.append((path, klass))
    return out


def naiti_pary(folder):
    """Список пар {id, kind, source, gotovyj}."""
    files = _excel_files(folder)
    by_dir = {}
    for path, klass in files:
        by_dir.setdefault(os.path.dirname(path), []).append((path, klass))

    pairs = []
    for directory, items in sorted(by_dir.items()):
        sources_o = [p for p, k in items if k == "source_oboi"]
        gotovye_o = [p for p, k in items if k == "gotovyj_oboi"]
        sources_y = [p for p, k in items if k == "source_yusuf"]
        gotovye_y = [p for p, k in items if k == "gotovyj_yusuf"]
        sources_v = [p for p, k in items if k == "source_vitebsk"]
        gotovye_v = [p for p, k in items if k == "gotovyj_vitebsk"]
        pairs.extend(_match(directory, "oboi", sources_o, gotovye_o))
        pairs.extend(_match(directory, "yusuf", sources_y, gotovye_y))
        pairs.extend(_match(directory, "vitebsk", sources_v, gotovye_v))
    return pairs


def _match(directory, kind, sources, gotovye):
    pairs = []
    used_g = set()
    for src in sources:
        best = None
        best_score = -1
        best_mtime = -1
        for got in gotovye:
            if got in used_g:
                continue
            score = _stem_score(os.path.basename(src), os.path.basename(got))
            if len(sources) == 1:
                score = max(score, 1.0)
            mt = _mtime(got)
            better = score > best_score + 1e-9
            same = abs(score - best_score) <= 1e-9 and mt >= best_mtime
            if better or same:
                best_score = score
                best_mtime = mt
                best = got
        if best is None or best_score < 0.15:
            continue
        used_g.add(best)
        src_name = os.path.basename(src)
        got_name = os.path.basename(best)
        pairs.append({
            "id": "%s|%s|%s" % (kind, src_name, got_name),
            "source_key": "%s|%s" % (kind, src_name),
            "kind": kind,
            "folder": directory,
            "source": src,
            "gotovyj": best,
        })
    return pairs


def _load_state(folder):
    path = os.path.join(folder, STATE_NAME)
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {"done": []}
    return {"done": []}


def _save_state(folder, state):
    path = os.path.join(folder, STATE_NAME)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)


def sravnit_oboi(out_path, got_path):
    """Сверка результата скрипта с эталоном. Возвращает список строк-расхождений."""
    from openpyxl import load_workbook

    def art_key(v):
        if isinstance(v, float) and v.is_integer():
            return str(int(v))
        return str(v).strip() if v is not None else ""

    out = load_workbook(out_path, data_only=True)["Обработанные данные"]
    got = load_workbook(got_path, data_only=True).worksheets[0]

    got_rows = []
    for r in range(1, got.max_row + 1):
        code = got.cell(r, 1).value
        if not (isinstance(code, str) and code.strip().upper().startswith("УТ")):
            continue
        got_rows.append({
            "code": str(code).strip(),
            "name": str(got.cell(r, 18).value).strip(),
            "art": art_key(got.cell(r, 19).value),
            "qty": str(got.cell(r, 36).value).strip(),
            "price": str(got.cell(r, 42).value).strip(),
            "sum_wo": str(got.cell(r, 49).value).strip(),
            "nds": str(got.cell(r, 68).value).strip(),
            "sum_w": str(got.cell(r, 76).value).strip(),
        })

    out_rows = []
    for r in range(2, out.max_row + 1):
        out_rows.append({
            "code": str(out.cell(r, 1).value or "").strip(),
            "name": str(out.cell(r, 3).value).strip(),
            "art": art_key(out.cell(r, 4).value),
            "qty": str(out.cell(r, 8).value).strip(),
            "price": str(out.cell(r, 9).value).strip(),
            "sum_wo": str(out.cell(r, 10).value).strip(),
            "nds": str(out.cell(r, 13).value).strip(),
            "sum_w": str(out.cell(r, 14).value).strip(),
        })

    diffs = []
    if len(out_rows) != len(got_rows):
        diffs.append("разное число строк: результат %s, эталон %s" % (
            len(out_rows), len(got_rows)))
        return diffs

    fields = ["code", "name", "art", "qty", "price", "sum_wo", "nds", "sum_w"]
    for i, (o, g) in enumerate(zip(out_rows, got_rows), 1):
        bad = [f for f in fields if o[f] != g[f]]
        if bad:
            diffs.append("#%s: " % i + "; ".join(
                "%s: %r != %r" % (f, o[f], g[f]) for f in bad))
    return diffs


def _arhivirovat(folder, pair, report_text):
    stamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    src_stem = os.path.splitext(os.path.basename(pair["source"]))[0]
    dest = os.path.join(folder, ARCHIVE_DIR, "%s_%s" % (stamp, src_stem[:40]))
    os.makedirs(dest, exist_ok=True)
    for key in ("source", "gotovyj"):
        src = pair[key]
        shutil.copy2(src, os.path.join(dest, os.path.basename(src)))
    with open(os.path.join(dest, "отчёт.txt"), "w", encoding="utf-8") as f:
        f.write(report_text)
    return dest


def obuchit_po_pare(pair, catalog_folder=None):
    """Разбирает одну пару и обновляет справочник. Возвращает dict-отчёт."""
    catalog_folder = catalog_folder or core.app_dir()
    kind = pair["kind"]
    source = pair["source"]
    gotovyj = pair["gotovyj"]
    report = {
        "id": pair["id"],
        "kind": kind,
        "source": os.path.basename(source),
        "gotovyj": os.path.basename(gotovyj),
        "added": 0,
        "updated": 0,
        "rows": 0,
        "diffs": [],
        "ok": False,
        "message": "",
    }

    log("", "info")
    log("Пара: %s" % pair["id"], "head")
    log("  исходник : %s" % source)
    log("  готовый  : %s" % gotovyj)

    if kind == "vitebsk":
        report["message"] = (
            "Витебск: эталон сохранён в архив. "
            "Новые коллекции правятся в convert_nakladnaya.py (см. документацию)."
        )
        log(report["message"], "warn")
        report["ok"] = True
        return report

    if kind == "yusuf":
        profile, prof_path, n_learned = core.learn_yusuf_from_gotovyj(
            gotovyj, folder=catalog_folder, source_path=source)
        report["rows"] = n_learned
        report["added"] = n_learned
        tmp_out = os.path.join(os.path.dirname(source), "_проверка_юсуф.xlsx")
        try:
            rows, out_path = core.convert_yusuf(source, tmp_out, profile=profile)
            report["rows"] = rows
            diffs = core.sravnit_yusuf(out_path, gotovyj)
            report["diffs"] = diffs
            if diffs:
                report["message"] = (
                    "Формат Юсуфа сохранён, но сверка с эталоном дала %s расхождений"
                    % len(diffs))
                log(report["message"], "warn")
                for line in diffs[:15]:
                    log("    " + line, "warn")
            else:
                report["message"] = (
                    "Формат Юсуфа сохранён (%s), сверка с эталоном: 0 расхождений"
                    % os.path.basename(prof_path))
                log(report["message"], "ok")
            report["ok"] = True
        except Exception as exc:
            report["message"] = "Профиль Юсуфа сохранён, проверка исходника не удалась: %s" % exc
            log(report["message"], "warn")
            report["ok"] = True
        finally:
            if os.path.exists(tmp_out):
                try:
                    os.remove(tmp_out)
                except OSError:
                    pass
        return report

    before = core.load_oboi_catalog(catalog_folder)
    core.learn_oboi_catalog_from_gotovyj(gotovyj, folder=catalog_folder)
    after = core.load_oboi_catalog(catalog_folder)
    added = [k for k in after if k not in before]
    updated = [k for k in after if k in before and before[k] != after[k]]
    report["added"] = len(added)
    report["updated"] = len(updated)
    if added:
        log("Новые артикулы: " + ", ".join(added[:12]) + ("…" if len(added) > 12 else ""), "ok")
    if updated:
        log("Обновлены артикулы: " + ", ".join(updated[:12]), "ok")

    tmp_out = os.path.join(os.path.dirname(source), "_проверка_агента.xlsx")
    try:
        rows, out_path = core.convert_oboi(source, tmp_out, catalog=after)
        report["rows"] = rows
        diffs = sravnit_oboi(out_path, gotovyj)
        report["diffs"] = diffs
        if diffs:
            report["message"] = "Справочник обновлён, но сверка с эталоном дала %s расхождений" % len(diffs)
            log(report["message"], "warn")
            for line in diffs[:15]:
                log("    " + line, "warn")
            report["ok"] = True
        else:
            report["message"] = "Справочник обновлён, сверка с эталоном: 0 расхождений"
            log(report["message"], "ok")
            report["ok"] = True
    except Exception as exc:
        report["message"] = "Справочник обновлён, проверка исходника не удалась: %s" % exc
        log(report["message"], "warn")
        report["ok"] = True
    finally:
        if os.path.exists(tmp_out):
            try:
                os.remove(tmp_out)
            except OSError:
                pass

    return report


def bump_app_version(new_version=None):
    """Поднимает номер версии на диске (version.txt)."""
    old = core.current_version()
    version = new_version or core.bump_patch(old)
    core.save_version(version)
    return old, version


def _polozhit_spravochniki(dest, catalog_folder):
    """Кладёт рядом с программой справочники, без которых обучение не уедет к сотруднику."""
    os.makedirs(dest, exist_ok=True)
    cat_src = os.path.join(catalog_folder, core.OBOI_CATALOG_FILENAME)
    if not os.path.exists(cat_src):
        cat_src = core.oboi_catalog_path(catalog_folder)
    if os.path.exists(cat_src):
        shutil.copy2(cat_src, os.path.join(dest, core.OBOI_CATALOG_FILENAME))
    else:
        cat = core.load_oboi_catalog(catalog_folder)
        if cat:
            core.save_oboi_catalog(cat, dest)

    yusuf_src = os.path.join(catalog_folder, core.YUSUF_PROFILE_FILENAME)
    if not os.path.exists(yusuf_src):
        yusuf_src = core.yusuf_profile_path(catalog_folder)
    if os.path.exists(yusuf_src):
        shutil.copy2(yusuf_src, os.path.join(dest, core.YUSUF_PROFILE_FILENAME))
    else:
        core.save_yusuf_profile(core.load_yusuf_profile(catalog_folder), dest)


def sobrat_paket_dlya_sotrudnika(catalog_folder, version, exe_src):
    """Один .exe со вшитыми справочниками — его отправляют сотруднику."""
    catalog_folder = os.path.abspath(catalog_folder)
    dest = os.path.join(catalog_folder, "Накладные_%s.exe" % version)

    oboi_path = os.path.join(catalog_folder, core.OBOI_CATALOG_FILENAME)
    if os.path.exists(oboi_path):
        with open(oboi_path, "r", encoding="utf-8") as f:
            oboi = json.load(f)
    else:
        oboi = core.load_oboi_catalog(catalog_folder)

    yusuf = core.load_yusuf_profile(catalog_folder)
    core.vshit_overlay_v_exe(exe_src, dest, {
        "version": version,
        "oboi_catalog": oboi or {},
        "yusuf_profile": yusuf or {},
    })

    pack_dir = os.path.join(catalog_folder, "релизы", "Накладные_%s" % version)
    if os.path.isdir(pack_dir):
        shutil.rmtree(pack_dir)
    os.makedirs(pack_dir, exist_ok=True)
    pack_exe = os.path.join(pack_dir, "Накладные.exe")
    shutil.copy2(dest, pack_exe)

    log("Один файл для сотрудника: %s" % dest, "ok")
    return {
        "version": version,
        "exe": dest,
        "pack_dir": pack_dir,
        "pack_exe": pack_exe,
    }


def vypustit_novyj_exe(catalog_folder, version):
    """
    Собирает один Накладные_версия.exe со справочниками внутри.
    Его можно отправить сотруднику. Старый открытый .exe не трогаем.
    """
    catalog_folder = os.path.abspath(catalog_folder)

    if getattr(sys, "frozen", False):
        src = os.path.abspath(sys.executable)
    else:
        log("Сейчас запуск не из .exe — новый файл для сотрудника не собран. "
            "Запустите Накладные.exe (не Python).", "warn")
        return None

    _polozhit_spravochniki(catalog_folder, catalog_folder)
    core.save_version(version, catalog_folder)
    pack = sobrat_paket_dlya_sotrudnika(catalog_folder, version, src)

    name = os.path.basename(pack["exe"])
    bat = os.path.join(catalog_folder, "Открыть_Накладные_%s.bat" % version)
    with open(bat, "w", encoding="utf-8") as f:
        f.write("@echo off\r\n")
        f.write("start \"\" \"%~dp0%s\"\r\n" % name)

    log("Новая программа: %s" % pack["exe"], "ok")
    return {
        "version": version,
        "exe": pack["exe"],
        "pack_dir": pack["pack_dir"],
        "pack_exe": pack["pack_exe"],
        "bat": bat,
    }


def zapisat_obnovlenie(reports, catalog_folder=None, release_info=None):
    """Пишет рядом с программой файл ЧТО_ОБНОВИЛОСЬ.txt."""
    catalog_folder = catalog_folder or core.app_dir()
    cat_path = os.path.join(catalog_folder, core.OBOI_CATALOG_FILENAME)
    if not os.path.exists(cat_path):
        cat_path = core.oboi_catalog_path(catalog_folder)
    added = sum(r.get("added", 0) for r in reports)
    updated = sum(r.get("updated", 0) for r in reports)
    lines = [
        "ЧТО ОБНОВИЛОСЬ ПОСЛЕ КНОПКИ «Обучить»",
        "",
    ]
    if release_info and release_info.get("exe"):
        lines += [
            "Один файл — отправьте сотруднику:",
            release_info["exe"],
            "",
            "Правила обработки уже внутри. Других файлов не нужно.",
            "",
            "Закройте старое окно и откройте этот файл у себя.",
            "Или двойной щелчок: %s" % os.path.basename(release_info.get("bat") or ""),
            "",
        ]
    else:
        lines += [
            "Файл .exe не скопирован (запуск был не из программы, а из Python).",
            "",
        ]
    lines += [
        "Справочник (новые позиции):",
        cat_path,
        "",
        "Добавлено артикулов: %s" % added,
        "Обновлено артикулов: %s" % updated,
        "",
        "Юсуф (packing list): %s" % os.path.join(
            catalog_folder, core.YUSUF_PROFILE_FILENAME),
        "",
        "Время: %s" % datetime.datetime.now().strftime("%d.%m.%Y %H:%M"),
        "",
    ]
    for r in reports:
        lines.append("- %s + %s: +%s новых, %s обновлено. %s" % (
            r["source"], r["gotovyj"], r["added"], r["updated"], r["message"]))
    path = os.path.join(catalog_folder, "ЧТО_ОБНОВИЛОСЬ.txt")
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
    log("Памятка: %s" % path, "muted")
    return path, cat_path


def process_inbox(folder=None, force=False, catalog_folder=None):
    """Находит новые пары в папке НовыеНакладные и обучает справочник."""
    folder = sozdat_papku_novye(folder)
    catalog_folder = catalog_folder or core.app_dir()
    started = datetime.datetime.now()
    itog = {
        "folder": folder,
        "done": [],
        "skipped": [],
        "errors": [],
        "rows": 0,
        "seconds": 0.0,
        "reports": [],
        "learned": True,
        "catalog_path": os.path.join(catalog_folder, core.OBOI_CATALOG_FILENAME),
    }

    log("Папка обучения: %s" % folder, "head")
    log("Справочник: %s" % core.oboi_catalog_path(catalog_folder))
    pairs = naiti_pary(folder)
    state = _load_state(folder)
    by_source = state.get("by_source") or {}

    if not pairs:
        log("Пар «исходник + готовый» не найдено.", "warn")
        log("Положите два файла в эту папку (см. %s)." % INSTRUCTION_NAME)
        itog["seconds"] = (datetime.datetime.now() - started).total_seconds()
        return itog

    log("Найдено пар: %s" % len(pairs))
    for pair in pairs:
        key = pair.get("source_key") or pair["id"]
        sig = _file_sig(pair["gotovyj"])
        prev = by_source.get(key)
        if not force and prev and prev.get("sig") == sig:
            itog["skipped"].append(key)
            log("Уже обучен (готовый не менялся): %s" % os.path.basename(pair["source"]), "muted")
            continue
        if prev and prev.get("sig") != sig:
            log("Исходник тот же, готовый файл другой — "
                "перезаписываю правила (%s)." % pair["kind"], "ok")
        try:
            report = obuchit_po_pare(pair, catalog_folder=catalog_folder)
            text = json.dumps(report, ensure_ascii=False, indent=2)
            arch = _arhivirovat(folder, pair, text)
            log("Архив пары: %s" % arch, "muted")
            itog["reports"].append(report)
            out_name = (
                core.YUSUF_PROFILE_FILENAME if report["kind"] == "yusuf"
                else core.OBOI_CATALOG_FILENAME
            )
            itog["done"].append({
                "src": report["source"],
                "out": out_name,
                "rows": report["rows"],
                "type": "learn-" + report["kind"],
            })
            itog["rows"] += report["rows"]
            by_source[key] = {
                "gotovyj": os.path.basename(pair["gotovyj"]),
                "sig": sig,
                "kind": pair["kind"],
                "learned_at": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }
        except Exception as exc:
            itog["errors"].append((pair["id"], str(exc)))
            log("ОШИБКА пары %s: %s" % (pair["id"], exc), "err")

    state["by_source"] = by_source
    state["done"] = sorted(by_source.keys())
    state["updated_at"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    _save_state(folder, state)

    if itog["reports"]:
        old_v, new_v = bump_app_version()
        log("Версия программы: %s → %s" % (old_v, new_v), "ok")
        release_info = None
        try:
            release_info = vypustit_novyj_exe(catalog_folder, new_v)
        except Exception as exc:
            log("Не удалось скопировать .exe: %s" % exc, "err")
            itog["errors"].append(("exe", str(exc)))
        note, cat_path = zapisat_obnovlenie(
            itog["reports"], catalog_folder=catalog_folder, release_info=release_info
        )
        itog["catalog_path"] = cat_path
        itog["note_path"] = note
        itog["version"] = new_v
        if release_info:
            itog["new_exe"] = release_info["exe"]
            itog["pack_dir"] = release_info["pack_dir"]
            itog["done"].append({
                "src": "для сотрудника",
                "out": release_info["exe"],
                "rows": 0,
                "type": "exe",
            })

    itog["seconds"] = (datetime.datetime.now() - started).total_seconds()
    return itog


def watch_inbox(folder=None, poll=4, catalog_folder=None):
    folder = sozdat_papku_novye(folder)
    log("Слежу за папкой: %s" % folder, "head")
    log("Положите пару файлов и подождите несколько секунд. Ctrl+C — стоп.")
    try:
        while True:
            process_inbox(folder, catalog_folder=catalog_folder)
            time.sleep(poll)
    except KeyboardInterrupt:
        log("Наблюдение остановлено.", "muted")
        return 0


def main(argv=None):
    parser = argparse.ArgumentParser(description="Агент обучения по папке НовыеНакладные")
    parser.add_argument("--setup", action="store_true", help="Создать папку НовыеНакладные")
    parser.add_argument("--once", action="store_true", help="Разобрать новые пары один раз")
    parser.add_argument("--watch", action="store_true", help="Следить за папкой")
    parser.add_argument("--folder", default=None, help="Путь к папке (по умолчанию C:\\НовыеНакладные)")
    parser.add_argument("--force", action="store_true", help="Обучить пары повторно")
    parser.add_argument("--poll", type=int, default=4)
    args = parser.parse_args(argv)

    folder = args.folder or default_novye_dir()
    if args.setup or (not args.once and not args.watch):
        path = sozdat_papku_novye(folder)
        log("Папка готова: %s" % path, "ok")
        log("Инструкция: %s" % os.path.join(path, INSTRUCTION_NAME))
        if not args.once and not args.watch:
            return 0
    if args.watch:
        return watch_inbox(folder, poll=args.poll) or 0
    itog = process_inbox(folder, force=args.force)
    log()
    if itog["errors"]:
        log("Завершено с ошибками.", "err")
        return 1
    if itog["done"]:
        log("Обучение завершено. Обработано пар: %s" % len(itog["done"]), "ok")
        return 0
    log("Новых пар нет.", "warn")
    return 0


if __name__ == "__main__":
    sys.exit(main())
