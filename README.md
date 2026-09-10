# Накладные: Витебск + Обои

Скрипт переводит накладные поставщиков в формат для учётной программы.

| Тип | Вход | Результат |
|---|---|---|
| Витебские ковры | `.xls` фактура | `*_VITEBSK_Обработано.xlsx` |
| Обои (УПД ВВП) | `.xlsx` УПД с длинными именами | `*_ОБОИ_Обработано.xlsx` |

Полная история решений и правила разбора — в [`VtebskieKovri_ДОКУМЕНТАЦИЯ.md`](VtebskieKovri_ДОКУМЕНТАЦИЯ.md).

## Состав репозитория

| Файл | Назначение |
|---|---|
| `convert_nakladnaya.py` | Основной скрипт (ковры + обои) |
| `Обработать_накладные.bat` | Запуск двойным щелчком на Windows |
| `oboi_catalog.json` | Справочник артикул → код УТ + короткое имя |
| `compare_oboi.py` | Сверка результата с эталоном «готовый» |
| `samples/` | Исходник и эталон УПД 04.09.26 |
| `requirements.txt` | Зависимости Python |

## Установка

```bash
pip install -r requirements.txt
```

На Windows в рабочей папке (`C:\Накладные`) достаточно один раз:

```bat
pip install openpyxl xlrd
```

## Запуск

Двойной щелчок по `Обработать_накладные.bat` — обработает все **новые**
`.xls` / `.xlsx` в папке скрипта.

```bat
python convert_nakladnaya.py
python convert_nakladnaya.py "файл.xlsx"
python convert_nakladnaya.py "файл.xlsx" --force
python convert_nakladnaya.py --learn-oboi "ВВП УПД … готовый .xlsx"
```

## Обои: справочник

Короткие имена и коды `УТ…` берутся из `oboi_catalog.json` (ключ — артикул).
Эталон из `samples/` уже занесён в справочник (50 позиций).
Новый эталон «готовый» подключайте через `--learn-oboi`.

## Проверка на образцах

```bash
python convert_nakladnaya.py --force ^
  "samples/ВВП УПД 04.09.26 № 00УТ-011419 source.xlsx" ^
  samples/out_oboi.xlsx

python compare_oboi.py samples/out_oboi.xlsx "samples/ВВП УПД 04.09.26 готовый.xlsx"
```

Ожидаемый результат: `OK: 0 расхождений`.
