# Накладные: Витебск + Обои

Скрипт `convert_nakladnaya.py` обрабатывает два типа файлов:

| Тип | Вход | Результат |
|---|---|---|
| Витебские ковры | `.xls` фактура | `*_VITEBSK_Обработано.xlsx` |
| Обои (УПД ВВП) | `.xlsx` УПД с длинными именами | `*_ОБОИ_Обработано.xlsx` |

Подробности по коврам и обоям — в `VtebskieKovri_ДОКУМЕНТАЦИЯ.md`.

## Установка

```bash
pip install -r requirements.txt
```

На Windows в `C:\Накладные` достаточно один раз поставить `xlrd` и `openpyxl`.

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
Эталон `ВВП УПД 04.09.26 готовый .xlsx` уже зашит в справочник (50 позиций).
Новый эталон «готовый» подключайте через `--learn-oboi`.

Проверка на образцах:

```bash
python convert_nakladnaya.py --force "samples/ВВП УПД 04.09.26 № 00УТ-011419 source.xlsx" samples/out_oboi.xlsx
```

Сверка с эталоном: код, имя, артикул, количество, цена, суммы, НДС — **0 расхождений**.
