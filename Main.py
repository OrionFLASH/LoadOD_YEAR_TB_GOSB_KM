#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Скрипт обработки файлов ОД (версия v3.0).

Ключевые доработки:
- Приоритет ToDo: поддержан файл OrgUnit (CSV) из IN/OrgUnit.
- Нормализация ключей:
  - ИНН_12 -> 12 символов с лидирующими нулями;
  - ТН 10 -> 8 символов с лидирующими нулями.
- Расчет динамики по последнему КМ сохранен.
- Добавлена колонка "Кластер" через merge с OrgUnit.
- Поддержан опциональный Config.json (если есть рядом со скриптом).
- Добавлено файловое логирование INFO/DEBUG в папку log.
"""

from __future__ import annotations

import glob
import json
import logging
import os
import re
import warnings
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from typing import Any

import pandas as pd
from openpyxl import Workbook

warnings.filterwarnings("ignore")

# ===== 1. БАЗОВЫЕ ПАРАМЕТРЫ =====
SCRIPT_DIR: str = os.path.dirname(os.path.abspath(__file__))
TIMESTAMP: str = datetime.now().strftime("%Y_%m_%d-%H_%M_%S")
LOG_STAMP: str = datetime.now().strftime("%Y%m%d_%H")
RUN_YEAR: str = datetime.now().strftime("%Y")
RUN_DAY_MONTH: str = datetime.now().strftime("%d-%m")
RUN_MONTH_DAY: str = datetime.now().strftime("%m-%d")


def load_config() -> dict[str, Any]:
    """Загружает настройки из Config.json (файл обязателен)."""
    config_path = os.path.join(SCRIPT_DIR, "Config.json")
    if not os.path.exists(config_path):
        raise FileNotFoundError("Не найден обязательный файл Config.json.")

    with open(config_path, "r", encoding="utf-8") as file:
        loaded = json.load(file)
    if not isinstance(loaded, dict):
        raise ValueError("Config.json должен содержать объект JSON.")
    return loaded


CONFIG: dict[str, Any] = load_config()
COLUMNS_TO_LOAD: list[str] = list(CONFIG["columns_to_load"])
GROUP_KEY: list[str] = list(CONFIG["group_key"])

INPUT_FOLDER: str = os.path.join(SCRIPT_DIR, str(CONFIG["paths"]["input_folder"]))
OUTPUT_FOLDER: str = os.path.join(SCRIPT_DIR, str(CONFIG["paths"]["output_folder"]))
LOG_FOLDER: str = os.path.join(SCRIPT_DIR, str(CONFIG["paths"]["log_folder"]))
ORGUNIT_FOLDER: str = os.path.join(SCRIPT_DIR, str(CONFIG["paths"]["orgunit_folder"]))

INPUT_FILE_GLOB: str = str(CONFIG["patterns"]["input_file_glob"])
ORGUNIT_FILE_GLOB: str = str(CONFIG["patterns"]["orgunit_file_glob"])
INPUT_SHEET: str = str(CONFIG["excel"]["input_sheet"])
MAX_WORKERS: int = int(CONFIG["processing"]["max_workers"])
XLSX_EXPORT_LIMIT: int = int(CONFIG["processing"]["xlsx_export_limit"])

OUTPUT_LAYOUT_BY_DATE: bool = bool(CONFIG["output"]["layout_by_date"])
OUTPUT_DATE_LAYOUT: str = str(CONFIG["output"]["date_layout"])
RAW_BASENAME: str = str(CONFIG["output"]["raw_basename"])
AGG_BASENAME: str = str(CONFIG["output"]["agg_basename"])
LAST_KM_BASENAME: str = str(CONFIG["output"]["last_km_basename"])
DYN_BASENAME: str = str(CONFIG["output"]["dyn_basename"])
FINAL_BASENAME: str = str(CONFIG["output"]["final_basename"])
STATS_BASENAME: str = str(CONFIG["output"]["stats_basename"])

os.makedirs(OUTPUT_FOLDER, exist_ok=True)
os.makedirs(LOG_FOLDER, exist_ok=True)

# Логи текущего запуска складываются в log/YYYY/MM-DD.
RUN_LOG_FOLDER: str = os.path.join(LOG_FOLDER, RUN_YEAR, RUN_MONTH_DAY)
os.makedirs(RUN_LOG_FOLDER, exist_ok=True)

# Все выходные артефакты текущего запуска складываются в OUT по шаблону из Config.json.
if OUTPUT_LAYOUT_BY_DATE and OUTPUT_DATE_LAYOUT == "YYYY/DD-MM":
    RUN_OUTPUT_FOLDER: str = os.path.join(OUTPUT_FOLDER, RUN_YEAR, RUN_DAY_MONTH)
else:
    RUN_OUTPUT_FOLDER = OUTPUT_FOLDER
os.makedirs(RUN_OUTPUT_FOLDER, exist_ok=True)


def setup_logger() -> logging.Logger:
    """Настраивает файловое логирование INFO/DEBUG по требуемому шаблону."""
    logger = logging.getLogger("loadod")
    logger.setLevel(logging.DEBUG)
    logger.handlers.clear()
    logger.propagate = False

    info_path = os.path.join(RUN_LOG_FOLDER, f"INFO_parser_{LOG_STAMP}.log")
    debug_path = os.path.join(RUN_LOG_FOLDER, f"DEBUG_parser_{LOG_STAMP}.log")

    info_handler = logging.FileHandler(info_path, encoding="utf-8")
    info_handler.setLevel(logging.INFO)
    info_handler.setFormatter(
        logging.Formatter("%(asctime)s - [%(levelname)s] - %(message)s", "%Y-%m-%d %H:%M:%S")
    )

    class DefaultDebugContextFilter(logging.Filter):
        """Добавляет значения по умолчанию для class_name и def_name."""

        def filter(self, record: logging.LogRecord) -> bool:
            if not hasattr(record, "class_name"):
                record.class_name = "-"
            if not hasattr(record, "def_name"):
                record.def_name = "-"
            return True

    debug_handler = logging.FileHandler(debug_path, encoding="utf-8")
    debug_handler.setLevel(logging.DEBUG)
    # Формат строго соответствует пользовательскому требованию.
    debug_handler.setFormatter(
        logging.Formatter(
            "%(asctime)s - [%(levelname)s] - %(message)s [class: %(class_name)s | def: %(def_name)s]",
            "%Y-%m-%d %H:%M:%S",
        )
    )

    logger.addHandler(info_handler)
    logger.addFilter(DefaultDebugContextFilter())
    logger.addHandler(debug_handler)
    return logger


LOGGER: logging.Logger = setup_logger()


def log_debug(message: str, class_name: str = "-", def_name: str = "-") -> None:
    """Единая точка DEBUG-логирования с обязательными полями class/def."""
    LOGGER.debug(message, extra={"class_name": class_name, "def_name": def_name})


def clean_numeric_string(value: Any) -> str:
    """Оставляет только цифры в значении, возвращает пустую строку для NaN."""
    if pd.isna(value):
        return ""
    raw = str(value).strip()
    digits = re.sub(r"\D", "", raw)
    return digits


def normalize_inn(value: Any) -> str:
    """Нормализация ИНН в 12-значный формат с лидирующими нулями."""
    return clean_numeric_string(value).zfill(12)[:12]


def normalize_tn10(value: Any) -> str:
    """Нормализация ТН 10 в 8-значный формат с лидирующими нулями."""
    return clean_numeric_string(value).zfill(8)[:8]


def normalize_code(value: Any) -> str:
    """Нормализация кодов ТБ/ГОСБ в строку без лишних пробелов."""
    if pd.isna(value):
        return ""
    return str(value).strip()


def save_table(df: pd.DataFrame, base_path_without_ext: str) -> int:
    """Сохранение таблицы: всегда CSV и XLSX, если строк меньше лимита."""
    rows = len(df)
    csv_path = f"{base_path_without_ext}.csv"
    df.to_csv(csv_path, index=False, sep=";", decimal=",", encoding="utf-8-sig")

    if rows < XLSX_EXPORT_LIMIT:
        xlsx_path = f"{base_path_without_ext}.xlsx"
        df.to_excel(xlsx_path, index=False, engine="openpyxl")

    return rows


def calc_growth_rate(row: pd.Series) -> float | None:
    """Темп прироста по правилам ToDo."""
    prev_val = row["Сумма ПРОШЛЫЙ ГОД ОД, тыс. руб."]
    curr_val = row["Сумма ТЕКУЩИЙ ГОД ОД, тыс. руб."]

    if pd.isna(prev_val) or pd.isna(curr_val):
        return None
    if prev_val == 0:
        if curr_val > 0:
            return 100.0
        if curr_val < 0:
            return -100.0
        return 0.0
    return (curr_val - prev_val) / abs(prev_val) * 100.0


def get_max_od_row(group: pd.DataFrame) -> pd.Series:
    """Возвращает запись с максимальным текущим ОД в группе."""
    if group["ТЕКУЩИЙ ГОД ОД, тыс. руб."].isna().all():
        return group.iloc[0]
    return group.loc[group["ТЕКУЩИЙ ГОД ОД, тыс. руб."].idxmax()]


def load_single_file(file_path: str) -> tuple[pd.DataFrame | None, dict[str, Any]]:
    """Загрузка одного Excel-файла с очисткой и нормализацией."""
    file_name = os.path.basename(file_path)
    start_time = datetime.now()

    file_stat: dict[str, Any] = {
        "Файл": file_name,
        "Статус": "OK",
        "Ошибка": "",
        "Строк исходных": 0,
        "Строк после фильтра": 0,
        "Удалено серая зона": 0,
        "Ошибок типов": 0,
        "Время загрузки (сек)": 0.0,
    }

    try:
        df = pd.read_excel(
            file_path,
            sheet_name=INPUT_SHEET,
            usecols=COLUMNS_TO_LOAD,
            engine="openpyxl",
        )
        file_stat["Строк исходных"] = len(df)
        df["ИмяФайла"] = file_name

        # Нормализация ключевых идентификаторов по приоритету ToDo.
        df["ИНН_12"] = df["ИНН_12"].apply(normalize_inn)
        df["ТН 10"] = df["ТН 10"].apply(normalize_tn10)
        df["Код ТБ"] = df["Код ТБ"].apply(normalize_code)
        df["Код ГОСБ"] = df["Код ГОСБ"].apply(normalize_code)
        df["ТБ"] = df["ТБ"].astype(str).str.strip()
        df["ГОСБ"] = df["ГОСБ"].astype(str).str.strip()
        df["КМ"] = df["КМ"].astype(str).str.strip()

        df["Дата загрузки"] = pd.to_datetime(df["Дата загрузки"], errors="coerce")
        df["ПРОШЛЫЙ ГОД ОД, тыс. руб."] = pd.to_numeric(
            df["ПРОШЛЫЙ ГОД ОД, тыс. руб."], errors="coerce", downcast="float"
        )
        df["ТЕКУЩИЙ ГОД ОД, тыс. руб."] = pd.to_numeric(
            df["ТЕКУЩИЙ ГОД ОД, тыс. руб."], errors="coerce", downcast="float"
        )

        rows_before_filter = len(df)
        df = df[df["КМ"] != "Серая зона"].copy()
        file_stat["Удалено серая зона"] = rows_before_filter - len(df)
        file_stat["Строк после фильтра"] = len(df)
        file_stat["Ошибок типов"] = (
            df["Дата загрузки"].isna().sum()
            + df["ПРОШЛЫЙ ГОД ОД, тыс. руб."].isna().sum()
            + df["ТЕКУЩИЙ ГОД ОД, тыс. руб."].isna().sum()
        )

        elapsed = (datetime.now() - start_time).total_seconds()
        file_stat["Время загрузки (сек)"] = round(elapsed, 2)
        log_debug(f"Файл загружен успешно: {file_name}", def_name="load_single_file")
        return df, file_stat
    except Exception as exc:
        file_stat["Статус"] = "ОШИБКА"
        file_stat["Ошибка"] = str(exc)
        file_stat["Время загрузки (сек)"] = round((datetime.now() - start_time).total_seconds(), 2)
        LOGGER.error("Ошибка загрузки файла %s: %s", file_name, exc)
        log_debug(f"Ошибка загрузки файла {file_name}: {exc}", def_name="load_single_file")
        return None, file_stat


def load_orgunit_mapping() -> pd.DataFrame:
    """Загружает первый CSV из IN/OrgUnit и возвращает маппинг ГОСБ -> Кластер."""
    if not os.path.exists(ORGUNIT_FOLDER):
        LOGGER.warning("Папка OrgUnit не найдена: %s", ORGUNIT_FOLDER)
        return pd.DataFrame(columns=["Код ГОСБ", "Кластер"])

    csv_files = sorted(glob.glob(os.path.join(ORGUNIT_FOLDER, ORGUNIT_FILE_GLOB)))
    if not csv_files:
        LOGGER.warning("CSV-файлы в OrgUnit не найдены: %s", ORGUNIT_FOLDER)
        return pd.DataFrame(columns=["Код ГОСБ", "Кластер"])

    org_path = csv_files[0]
    LOGGER.info("Загрузка OrgUnit: %s", os.path.basename(org_path))
    log_debug(f"Читаем OrgUnit файл: {org_path}", def_name="load_orgunit_mapping")

    # Пытаемся прочитать как ';', если не подходит - читаем как ','.
    org_df = pd.read_csv(org_path, sep=";", dtype=str, encoding="utf-8", engine="python")
    if "GOSB_CODE" not in org_df.columns or "CLUSTER" not in org_df.columns:
        org_df = pd.read_csv(org_path, sep=",", dtype=str, encoding="utf-8", engine="python")

    if "GOSB_CODE" not in org_df.columns or "CLUSTER" not in org_df.columns:
        raise ValueError(
            "Файл OrgUnit должен содержать колонки GOSB_CODE и CLUSTER."
        )

    org_df = org_df[["GOSB_CODE", "CLUSTER"]].copy()
    org_df["Код ГОСБ"] = org_df["GOSB_CODE"].apply(normalize_code)
    org_df["Кластер"] = org_df["CLUSTER"].astype(str).str.strip()
    org_df = org_df[["Код ГОСБ", "Кластер"]].drop_duplicates()
    return org_df


def auto_fit_worksheet(ws: Any) -> None:
    """Подгоняет ширину колонок листа Excel."""
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            cell_len = len(str(cell.value)) if cell.value is not None else 0
            if cell_len > max_length:
                max_length = cell_len
        ws.column_dimensions[column].width = min(max_length + 2, 60)


def main() -> None:
    """Основной сценарий обработки данных."""
    total_start = datetime.now()
    LOGGER.info("Старт обработки ОД")
    log_debug("Старт обработки ОД", def_name="main")

    print("=" * 70)
    print("НАЧАЛО ОБРАБОТКИ (ВЕРСИЯ v3.0)")
    print("=" * 70)
    print(f"Время запуска:         {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"Рабочая папка:         {SCRIPT_DIR}")
    print(f"Входная папка (IN):    {INPUT_FOLDER}")
    print(f"Папка OrgUnit:         {ORGUNIT_FOLDER}")
    print(f"Выходная папка (OUT):  {RUN_OUTPUT_FOLDER}")
    print(f"Папка логов:           {LOG_FOLDER}")
    print(f"Таймштамп:             {TIMESTAMP}")
    print(f"Потоков загрузки:      {MAX_WORKERS}")

    if not os.path.exists(INPUT_FOLDER):
        raise FileNotFoundError("Папка IN не существует.")

    file_list = sorted(glob.glob(os.path.join(INPUT_FOLDER, INPUT_FILE_GLOB)))
    file_list = [f for f in file_list if not os.path.basename(f).startswith("~$")]
    if not file_list:
        raise FileNotFoundError("В папке IN нет XLSX-файлов.")

    print(f"\nНайдено XLSX файлов: {len(file_list)}")
    for idx, fp in enumerate(file_list, start=1):
        print(f"  [{idx:2d}] {os.path.basename(fp)}")

    dfs: list[pd.DataFrame] = []
    stats: list[dict[str, Any]] = []

    print("\n" + "=" * 70)
    print(f"ЗАГРУЗКА ФАЙЛОВ (параллельно, {MAX_WORKERS} потоков)")
    print("=" * 70)

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        future_to_file = {executor.submit(load_single_file, path): path for path in file_list}
        completed = 0
        for future in as_completed(future_to_file):
            file_name = os.path.basename(future_to_file[future])
            completed += 1
            df_part, file_stat = future.result()
            stats.append(file_stat)
            if df_part is not None:
                dfs.append(df_part)
                print(
                    f"[{completed:2d}/{len(file_list)}] ✓ {file_name[:45]:<45} | "
                    f"{len(df_part):>8,} строк"
                )
            else:
                print(f"[{completed:2d}/{len(file_list)}] ✗ {file_name}: {file_stat['Ошибка'][:55]}")

    if not dfs:
        raise RuntimeError("Нет успешно загруженных файлов.")

    df_combined = pd.concat(dfs, ignore_index=True, copy=False)
    del dfs

    total_grey_removed = sum(s["Удалено серая зона"] for s in stats if s["Статус"] == "OK")
    errors_date = int(df_combined["Дата загрузки"].isna().sum())
    errors_prev = int(df_combined["ПРОШЛЫЙ ГОД ОД, тыс. руб."].isna().sum())
    errors_curr = int(df_combined["ТЕКУЩИЙ ГОД ОД, тыс. руб."].isna().sum())
    total_errors = errors_date + errors_prev + errors_curr

    output_raw_csv = os.path.join(RUN_OUTPUT_FOLDER, f"{RAW_BASENAME}_{TIMESTAMP}.csv")
    output_agg_csv = os.path.join(RUN_OUTPUT_FOLDER, f"{AGG_BASENAME}_{TIMESTAMP}.csv")
    output_last_km_base = os.path.join(RUN_OUTPUT_FOLDER, f"{LAST_KM_BASENAME}_{TIMESTAMP}")
    output_dyn_base = os.path.join(RUN_OUTPUT_FOLDER, f"{DYN_BASENAME}_{TIMESTAMP}")
    output_final_base = os.path.join(RUN_OUTPUT_FOLDER, f"{FINAL_BASENAME}_{TIMESTAMP}")
    output_stats_xlsx = os.path.join(RUN_OUTPUT_FOLDER, f"{STATS_BASENAME}_{TIMESTAMP}.xlsx")

    df_combined.to_csv(output_raw_csv, index=False, sep=";", decimal=",", encoding="utf-8-sig")

    agg_df = df_combined.groupby(GROUP_KEY, as_index=False, dropna=False, sort=False).agg(
        {
            "ПРОШЛЫЙ ГОД ОД, тыс. руб.": "sum",
            "ТЕКУЩИЙ ГОД ОД, тыс. руб.": "sum",
        }
    )
    agg_df.rename(
        columns={
            "ПРОШЛЫЙ ГОД ОД, тыс. руб.": "Сумма ПРОШЛЫЙ ГОД ОД, тыс. руб.",
            "ТЕКУЩИЙ ГОД ОД, тыс. руб.": "Сумма ТЕКУЩИЙ ГОД ОД, тыс. руб.",
        },
        inplace=True,
    )
    agg_df.to_csv(output_agg_csv, index=False, sep=";", decimal=",", encoding="utf-8-sig")

    df_for_last_km = df_combined[df_combined["Дата загрузки"].notna()].copy()
    grouped_by_date = (
        df_for_last_km.groupby(
            ["ИНН_12", "Код ТБ", "Код ГОСБ", "Дата загрузки"], as_index=False, dropna=False, sort=False
        )
        .apply(get_max_od_row, include_groups=False)
        .reset_index(drop=True)
    )

    last_dates = grouped_by_date.groupby(
        ["ИНН_12", "Код ТБ", "Код ГОСБ"], as_index=False, dropna=False, sort=False
    )["Дата загрузки"].max()

    last_km = pd.merge(
        last_dates,
        grouped_by_date,
        on=["ИНН_12", "Код ТБ", "Код ГОСБ", "Дата загрузки"],
        how="left",
    )

    last_km_result = last_km[
        ["ИНН_12", "Код ТБ", "Код ГОСБ", "ТБ", "ГОСБ", "КМ", "ТН 10", "Дата загрузки"]
    ].copy()
    last_km_result.rename(
        columns={
            "КМ": "Последний КМ",
            "ТН 10": "Последний ТН 10",
            "Дата загрузки": "Дата последнего закрепления",
        },
        inplace=True,
    )

    km_pivot = grouped_by_date.pivot_table(
        index=["ИНН_12", "Код ТБ", "Код ГОСБ"],
        columns="Дата загрузки",
        values="КМ",
        aggfunc="first",
        fill_value="",
    )
    km_dates = sorted(km_pivot.columns.dropna())
    km_pivot.columns = [f"КМ на {d.strftime('%d_%m_%Y')}" for d in km_dates]
    km_pivot = km_pivot.reset_index()

    last_km_with_dates = pd.merge(
        last_km_result,
        km_pivot,
        on=["ИНН_12", "Код ТБ", "Код ГОСБ"],
        how="left",
    )
    save_table(last_km_with_dates, output_last_km_base)

    last_km_key = last_km_result[["ИНН_12", "Код ТБ", "Код ГОСБ", "Последний КМ"]]
    df_with_last = pd.merge(
        df_combined,
        last_km_key,
        on=["ИНН_12", "Код ТБ", "Код ГОСБ"],
        how="left",
        suffixes=("", "_last"),
    )
    df_last_km_only = df_with_last[df_with_last["КМ"] == df_with_last["Последний КМ"]].copy()

    dyn_group = df_last_km_only.groupby(
        ["ТН 10", "КМ", "Код ТБ", "Код ГОСБ", "ТБ", "ГОСБ", "Дата загрузки"],
        as_index=False,
        dropna=False,
        sort=False,
    ).agg(
        {
            "ПРОШЛЫЙ ГОД ОД, тыс. руб.": "sum",
            "ТЕКУЩИЙ ГОД ОД, тыс. руб.": "sum",
        }
    )
    dyn_group.rename(
        columns={
            "ПРОШЛЫЙ ГОД ОД, тыс. руб.": "Сумма ПРОШЛЫЙ ГОД ОД, тыс. руб.",
            "ТЕКУЩИЙ ГОД ОД, тыс. руб.": "Сумма ТЕКУЩИЙ ГОД ОД, тыс. руб.",
        },
        inplace=True,
    )
    dyn_group["Прирост ОД, тыс. руб."] = (
        dyn_group["Сумма ТЕКУЩИЙ ГОД ОД, тыс. руб."] - dyn_group["Сумма ПРОШЛЫЙ ГОД ОД, тыс. руб."]
    )
    dyn_group["Темп прироста, %"] = dyn_group.apply(calc_growth_rate, axis=1)

    save_table(dyn_group, output_dyn_base)

    orgunit_map = load_orgunit_mapping()
    final_result = pd.merge(dyn_group, orgunit_map, on="Код ГОСБ", how="left")
    final_result["Кластер"] = final_result["Кластер"].fillna("НЕ НАЙДЕН")
    save_table(final_result, output_final_base)

    # ===== СТАТИСТИКА =====
    stats.sort(key=lambda x: x["Файл"])
    wb = Workbook()
    ws = wb.active
    ws.title = "Статистика"
    ws.append(
        [
            "Файл",
            "Статус",
            "Ошибка",
            "Строк исходных",
            "Строк после фильтра",
            "Удалено серая зона",
            "Ошибок типов",
            "Время загрузки (сек)",
        ]
    )
    for stat in stats:
        ws.append(
            [
                stat["Файл"],
                stat["Статус"],
                stat["Ошибка"],
                stat["Строк исходных"],
                stat["Строк после фильтра"],
                stat["Удалено серая зона"],
                stat["Ошибок типов"],
                stat["Время загрузки (сек)"],
            ]
        )

    ws.append([])
    ws.append(["ИТОГО:"])
    ws.append(["Время обработки", TIMESTAMP])
    ws.append(["Потоков", MAX_WORKERS])
    ws.append(["Всего XLSX файлов", len(file_list)])
    ws.append(["Успешно загружено", len([s for s in stats if s["Статус"] == "OK"])])
    ws.append(["С ошибками", len([s for s in stats if s["Статус"] == "ОШИБКА"])])
    ws.append(["Всего строк (без серой)", len(df_combined)])
    ws.append(["Удалено серой зоны", total_grey_removed])
    ws.append(["Агрегированных строк", len(agg_df)])
    ws.append(["Записей последнего КМ", len(last_km_result)])
    ws.append(["Строк в таблице КМ+даты", len(last_km_with_dates)])
    ws.append(["Строк в динамике", len(dyn_group)])
    ws.append(["Строк в финальном файле (с кластером)", len(final_result)])
    ws.append(["Ошибок типов", total_errors])
    ws.append(["OrgUnit записей", len(orgunit_map)])
    ws.append(
        [
            "OrgUnit непросопоставленных (НЕ НАЙДЕН)",
            int((final_result["Кластер"] == "НЕ НАЙДЕН").sum()),
        ]
    )
    ws.append(["Общее время (мин)", round((datetime.now() - total_start).total_seconds() / 60, 2)])
    auto_fit_worksheet(ws)
    wb.save(output_stats_xlsx)

    total_time = (datetime.now() - total_start).total_seconds()
    LOGGER.info("Обработка завершена успешно за %.2f сек", total_time)
    log_debug("Обработка завершена успешно", def_name="main")

    print("\n" + "=" * 70)
    print("ОБРАБОТКА ЗАВЕРШЕНА")
    print("=" * 70)
    print(f"⏱ ОБЩЕЕ ВРЕМЯ: {total_time:.1f} сек ({total_time / 60:.2f} мин)")
    print("\nВыходные файлы в OUT:")
    print(f"  1. {os.path.basename(output_raw_csv)}")
    print(f"  2. {os.path.basename(output_agg_csv)}")
    print(f"  3. {os.path.basename(output_last_km_base)}.csv[.xlsx]")
    print(f"  4. {os.path.basename(output_dyn_base)}.csv[.xlsx]")
    print(f"  5. {os.path.basename(output_final_base)}.csv[.xlsx]   <-- основной итог")
    print(f"  6. {os.path.basename(output_stats_xlsx)}")
    print("=" * 70)


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        LOGGER.exception("Критическая ошибка выполнения: %s", exc)
        print(f"\n✗ КРИТИЧЕСКАЯ ОШИБКА: {exc}")
        raise