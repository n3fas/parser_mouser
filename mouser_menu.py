import os
import sys
import json
import re
from pathlib import Path
from typing import List, Dict, Any

from dotenv import load_dotenv, find_dotenv

from mouser_cli import (
    _get_api_key,
    _is_valid_pn,
    fetch_by_partnumber,
    fetch_by_keyword,
    extract_first_part,
    transform_strict,
    write_csv,
    print_table,
    write_xlsx,
)
from db_cache import get_cached_part, save_cached_part

load_dotenv(find_dotenv(), override=False)

def _ask(prompt: str) -> str:
    try:
        return input(prompt).strip()
    except (EOFError, KeyboardInterrupt):
        print("\nВыход.")
        sys.exit(0)


def _split_pns(s: str) -> List[str]:
    # Регулярка ищет текст в двойных кавычках, в одинарных кавычках, либо текст без пробелов и запятых
    matches = re.finditer(r'"([^"]+)"|\'([^\']+)\'|([^,\s]+)', s)
    parts = []
    for m in matches:
        val = m.group(1) or m.group(2) or m.group(3)
        if val and val.strip():
            parts.append(val.strip())
    return parts


def _save_json(rows: List[Dict[str, Any]], out_path: str) -> None:
    p = Path(out_path)
    p.parent.mkdir(parents=True, exist_ok=True)
    with p.open("w", encoding="utf-8") as f:
        json.dump(rows, f, ensure_ascii=False, indent=2)


def _save_raw(data: Dict[str, Any], save_dir: Path, name: str) -> None:
    save_dir.mkdir(parents=True, exist_ok=True)
    safe = name.replace("/", "_")
    with (save_dir / f"{safe}.json").open("w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def _lookup_one(pn: str, api_key: str, retries: int = 3, timeout: int = 20,
                save_raw_dir: Path | None = None, force_update: bool = False) -> dict:
    
    if not force_update:
        cached = get_cached_part(pn)
        if cached:
            print(f"  ⚡ Найдено в локальной БД: {pn}")
            return cached

    empty_row = {
        "Запрошенный партномер": pn,
        "Партномер Mouser": "-",
        "Партномер производителя": "-",
        "Производитель": "-",
        "Категория (EN)": "-",
        "Категория": "-",
        "Описание": "-",
        "Ссылка на фото": "-",
        "Даташит": "-"
    }

    if not _is_valid_pn(pn):
        print(f"  ⚠️  Пропускаю некорректный ввод: {pn!r}")
        empty_row["Описание"] = "Invalid PN"
        return empty_row

    try:
        data_pn = fetch_by_partnumber(pn, api_key, retries, timeout)
    except Exception as e:
        print(f"  ❌ Ошибка запроса (partnumber): {e}")
        empty_row["Описание"] = f"Error: {e}"
        return empty_row

    part = extract_first_part(data_pn)
    used_keyword = False

    if not part:
        try:
            data_kw = fetch_by_keyword(pn, api_key, retries, timeout)
        except Exception as e:
            print(f"  ❌ Ошибка запроса (Формат Excel.xls устарел и не поддерживается, пожалуйста, используйте формат Excel.xlsx): {e}")
            empty_row["Описание"] = f"Error: {e}"
            return empty_row
        part = extract_first_part(data_kw)
        used_keyword = True
        if save_raw_dir:
            _save_raw(data_kw, save_raw_dir, f"{pn}__keyword")
    else:
        if save_raw_dir:
            _save_raw(data_pn, save_raw_dir, f"{pn}__partnumber")

    if not part:
        print("  ⚠️  Ничего не найдено.")
        empty_row["Описание"] = "Ничего не найдено"
        save_cached_part(pn, empty_row)
        return empty_row

    row = transform_strict(part, pn)
    
    # Replace None values with dashes in the resulting row
    for k, v in row.items():
        if v is None:
            row[k] = "-"
            
    save_cached_part(pn, row)
    print("  ✓ Готово.")
    return row


try:
    from tqdm import tqdm
except ImportError:
    def tqdm(iterable, **kwargs):
        return iterable

def _process_excel_file(filepath: str, api_key: str, save_raw_dir: Path | None):
    p = Path(filepath)
    if not p.exists():
        print(f"  ❌ Файл {filepath} не найден.")
        return

    try:
        wb = load_workbook(filepath, data_only=True)
        ws = wb.active
    except Exception as e:
        print(f"  ❌ Ошибка чтения Excel: {e}")
        return

    part_col_idx = None
    start_row = None

    for row_idx, row in enumerate(ws.iter_rows(max_row=50), start=1):
        for col_idx, cell in enumerate(row, start=1):
            if cell.value and isinstance(cell.value, str):
                val_clean = cell.value.strip().lower().replace(" ", "").replace(".", "")
                if val_clean in ("partno", "partnumber", "part#", "pn"):
                    part_col_idx = col_idx
                    start_row = row_idx + 1
                    break
        if part_col_idx:
            break

    if not part_col_idx:
        print("  ❌ Столбец с партномером ('Part no.', 'PART NUMBER' и т.д.) не найден.")
        return

    print(f"  ✓ Найден целевой столбец (столбец {part_col_idx}, начиная со строки {start_row}).")

    pns = []
    for row_idx in range(start_row, ws.max_row + 1):
        val = ws.cell(row=row_idx, column=part_col_idx).value
        if val is None or str(val).strip() == "":
            break
        pns.append(str(val).strip())

    if not pns:
        print("  ⚠️  Список парт-номеров пуст.")
        return

    print(f"  ✓ Найдено {len(pns)} парт-номеров.")
    all_results = []

    for pn in tqdm(pns, desc="Processing Excel rows"):
        row = _lookup_one(pn, api_key, retries=3, timeout=20, save_raw_dir=save_raw_dir)
        if row:
            all_results.append(row)

    if all_results:
        out_name = f"mouser_excel_result_{p.stem}.xlsx"
        try:
            write_xlsx(all_results, out_name)
            print(f"\n  ✓ Готово! Результаты сохранены в {out_name}")
        except Exception as e:
            print(f"  ❌ Ошибка сохранения итогового файла: {e}")
    else:
        print("\n  ⚠️  Нет успешных результатов для сохранения.")

def main():
    try:
        api_key = _get_api_key()
    except SystemExit:
        print("Создайте .env с MOUSER_API_KEY=... и запустите снова.")
        return

    results_map: Dict[str, Dict[str, Any]] = {}
    save_raw_dir: Path | None = None

    while True:
        print("\n============== Mouser Parser (интерактивный режим) ==============")
        print("1) Ввести парт-номер(а) и показать таблицу")
        print("2) Сохранить последние результаты в CSV")
        print("3) Сохранить последние результаты в JSON")
        print("4) Указать папку для сохранения сырых ответов API (RAW)")
        print("5) Сохранить последние результаты в XLSX")
        print("6) Обработать Excel-файл (ищет столбец 'Part no.', 'PART NUMBER' и т.д.)")
        print("0) Выход")
        choice = _ask("\nВыберите пункт меню: ")

        if choice == "0":
            print("Пока! 👋")
            break

        elif choice == "1":
            raw = _ask("Введите один или несколько парт-номеров (через пробел или запятую): ")
            pns = _split_pns(raw)
            if not pns:
                print("  ⚠️  Ничего не введено.")
                continue

            for pn in tqdm(pns, desc="Searching"):
                row = _lookup_one(pn, api_key, retries=3, timeout=20, save_raw_dir=save_raw_dir)
                if row:
                    key = pn.strip().lower()
                    results_map[key] = row

            if results_map:
                print("\nРезультаты:")
                print_table(
                    list(results_map.values()), 
                    display_cols=["Запрошенный партномер", "Партномер Mouser", "Производитель", "Категория", "Описание"]
                )
            else:
                print("  ⚠️  Пусто — нет валидных результатов.")

        elif choice == "2":
            if not results_map:
                print("  ⚠️  Нет данных для сохранения. Сначала выполните поиск (пункт 1).")
                continue
            out = _ask("Имя файла CSV (по умолчанию: mouser_results.csv): ") or "mouser_results.csv"
            try:
                write_csv(list(results_map.values()), out)
                print(f"  ✓ CSV сохранён → {out}")
            except Exception as e:
                print(f"  ❌ Ошибка при сохранении CSV: {e}")

        elif choice == "3":
            if not results_map:
                print("  ⚠️  Нет данных для сохранения. Сначала выполните поиск (пункт 1).")
                continue
            out = _ask("Имя файла JSON (по умолчанию: mouser_results.json): ") or "mouser_results.json"
            try:
                _save_json(list(results_map.values()), out)
                print(f"  ✓ JSON сохранён → {out}")
            except Exception as e:
                print(f"  ❌ Ошибка при сохранении JSON: {e}")

        elif choice == "4":
            path = _ask("Укажи путь к папке для RAW (пусто — отключить): ")
            if not path:
                save_raw_dir = None
                print("  RAW сохранение отключено.")
            else:
                p = Path(path)
                try:
                    p.mkdir(parents=True, exist_ok=True)
                    save_raw_dir = p
                    print(f"  ✓ RAW будут сохраняться в: {p}")
                except Exception as e:
                    print(f"  ❌ Не удалось создать папку: {e}")
                    
        elif choice == "5":
            if not results_map:
                print("  ⚠️  Нет данных для сохранения. Сначала выполните поиск (пункт 1).")
                continue
            out = _ask("Имя файла XLSX (по умолчанию: mouser_results.xlsx): ") or "mouser_results.xlsx"
            try:
                write_xlsx(list(results_map.values()), out)
            except Exception as e:
                print(f"  ❌ Ошибка при сохранении XLSX: {e}")
                
        elif choice == "6":
            filepath = _ask("Введите путь к Excel файлу: ")
            if not filepath: continue
            _process_excel_file(filepath, api_key, save_raw_dir)
            
        else:
            print("  ⚠️  Неверный выбор. Попробуйте ещё раз.")

if __name__ == "__main__":
    main()
