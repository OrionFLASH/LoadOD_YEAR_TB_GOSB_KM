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
import time
import warnings
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from typing import Any

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows

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
MAX_WORKERS_CONFIG: int = int(CONFIG["processing"]["max_workers"])
XLSX_EXPORT_LIMIT: int = int(CONFIG["processing"]["xlsx_export_limit"])

OUTPUT_LAYOUT_BY_DATE: bool = bool(CONFIG["output"]["layout_by_date"])
OUTPUT_DATE_LAYOUT: str = str(CONFIG["output"]["date_layout"])
RAW_BASENAME: str = str(CONFIG["output"]["raw_basename"])
AGG_BASENAME: str = str(CONFIG["output"]["agg_basename"])
LAST_KM_BASENAME: str = str(CONFIG["output"]["last_km_basename"])
DYN_BASENAME: str = str(CONFIG["output"]["dyn_basename"])
FINAL_BASENAME: str = str(CONFIG["output"]["final_basename"])
STATS_BASENAME: str = str(CONFIG["output"]["stats_basename"])
CONSOLE_VERBOSITY: str = str(CONFIG.get("runtime", {}).get("console_verbosity", "normal")).lower()
if CONSOLE_VERBOSITY not in {"quiet", "normal", "verbose"}:
    CONSOLE_VERBOSITY = "normal"

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


def _verbosity_value(level: str) -> int:
    """Преобразует текстовый уровень детализации в число."""
    mapping = {"quiet": 0, "normal": 1, "verbose": 2}
    return mapping[level]


def cprint(message: str, level: str = "normal") -> None:
    """Печатает сообщение в консоль согласно уровню детализации."""
    if _verbosity_value(CONSOLE_VERBOSITY) >= _verbosity_value(level):
        print(message)


def fmt_elapsed(seconds: float) -> str:
    """Форматирует длительность в читаемый вид."""
    return f"{seconds:.2f}с"


def normalize_gosb_code(series: pd.Series) -> pd.Series:
    """
    Приводит «Код ГОСБ» к одному строковому виду для merge с OrgUnit.

    В Excel колонка часто числовая: при чтении получается float и строка вида «4001.0»,
    тогда как в CSV справочника — «4001», и merge не находит совпадений.
    Целые числа (в т.ч. из float) переводим в строку без десятичной части.
    """
    num = pd.to_numeric(series, errors="coerce")
    text = series.astype("string").fillna("").str.strip()
    # Целые значения — единый формат «123», без «123.0»
    is_whole = num.notna() & (num % 1 == 0)
    out = text.copy()
    out.loc[is_whole] = num.loc[is_whole].astype(np.int64).astype(str)
    return out


def resolve_max_workers(config_value: int) -> tuple[int, str]:
    """
    Определяет число потоков:
    - если config_value > 0: используем заданное значение;
    - если config_value == 0: авто по числу логических ядер.
    """
    if config_value > 0:
        return config_value, "fixed"
    cpu_count = os.cpu_count() or 1
    return max(1, cpu_count), "auto"


def should_export_to_xlsx(row_count: int) -> bool:
    """Определяет, помещается ли таблица в XLSX по лимиту строк."""
    return row_count < XLSX_EXPORT_LIMIT


def apply_sheet_formatting(ws: Any, headers: list[str]) -> None:
    """Применяет форматирование: заголовок, freeze, фильтры, форматы данных."""
    if not headers:
        return
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{ws.max_row}"

    for col_idx, col_name in enumerate(headers, start=1):
        col_letter = get_column_letter(col_idx)
        name = str(col_name)
        is_amount = ("ПРОШЛЫЙ" in name) or ("ТЕКУЩИЙ" in name) or ("Прирост" in name)
        is_date = "Дата" in name
        is_percent = "Темп прироста" in name

        if is_amount:
            num_fmt = "#,##0.00"
        elif is_date:
            num_fmt = "DD.MM.YYYY"
        elif is_percent:
            num_fmt = "0.00%"
        else:
            num_fmt = None

        if num_fmt is not None:
            for row_idx in range(2, ws.max_row + 1):
                ws[f"{col_letter}{row_idx}"].number_format = num_fmt

    auto_fit_worksheet(ws)


def write_df_to_sheet(wb: Workbook, sheet_name: str, df: pd.DataFrame) -> None:
    """Записывает DataFrame на лист с базовым форматированием."""
    ws = wb.create_sheet(title=sheet_name[:31])
    export_df = df.copy()
    for col in export_df.columns:
        if "Темп прироста" in str(col):
            export_df[col] = pd.to_numeric(export_df[col], errors="coerce") / 100.0
    # Пакетная запись через dataframe_to_rows быстрее, чем построчный itertuples + append.
    for row in dataframe_to_rows(export_df, index=False, header=True):
        ws.append(list(row))
    apply_sheet_formatting(ws, list(export_df.columns))


def build_stats_frames(
    stats: list[dict[str, Any]],
    timestamp: str,
    workers: int,
    workers_mode: str,
    df_combined: pd.DataFrame,
    total_grey_removed: int,
    agg_df: pd.DataFrame,
    last_km_result: pd.DataFrame,
    last_km_with_dates: pd.DataFrame,
    dyn_group: pd.DataFrame,
    final_result: pd.DataFrame,
    total_errors: int,
    orgunit_map: pd.DataFrame,
    total_minutes: float,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Готовит таблицу статистики по файлам и таблицу итогов."""
    stats_files_df = pd.DataFrame(stats)
    summary_df = pd.DataFrame(
        [
            ("Время обработки", timestamp),
            ("Потоков (факт)", workers),
            ("Режим потоков", workers_mode),
            ("Всего XLSX файлов", len(stats)),
            ("Успешно загружено", len([s for s in stats if s["Статус"] == "OK"])),
            ("С ошибками", len([s for s in stats if s["Статус"] == "ОШИБКА"])),
            ("Всего строк (без серой)", len(df_combined)),
            ("Удалено серой зоны", total_grey_removed),
            ("Агрегированных строк", len(agg_df)),
            ("Записей последнего КМ", len(last_km_result)),
            ("Строк в таблице КМ+даты", len(last_km_with_dates)),
            ("Строк в динамике", len(dyn_group)),
            ("Строк в финальном файле (с кластером)", len(final_result)),
            ("Ошибок типов", total_errors),
            ("OrgUnit записей", len(orgunit_map)),
            ("OrgUnit непросопоставленных (НЕ НАЙДЕН)", int((final_result["Кластер"] == "НЕ НАЙДЕН").sum())),
            ("Общее время (мин)", total_minutes),
        ],
        columns=["Показатель", "Значение"],
    )
    return stats_files_df, summary_df


def load_single_file(file_path: str) -> tuple[pd.DataFrame | None, dict[str, Any]]:
    """Загрузка одного Excel-файла с очисткой и нормализацией."""
    file_name = os.path.basename(file_path)
    start_time_perf = time.perf_counter()

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

        # Векторная нормализация ключевых полей: заметно быстрее apply() на больших данных.
        df["ИНН_12"] = (
            df["ИНН_12"]
            .astype("string")
            .fillna("")
            .str.replace(r"\D", "", regex=True)
            .str.zfill(12)
            .str.slice(0, 12)
        )
        df["ТН 10"] = (
            df["ТН 10"]
            .astype("string")
            .fillna("")
            .str.replace(r"\D", "", regex=True)
            .str.zfill(8)
            .str.slice(0, 8)
        )
        df["Код ТБ"] = df["Код ТБ"].astype("string").fillna("").str.strip()
        df["Код ГОСБ"] = normalize_gosb_code(df["Код ГОСБ"])
        df["ТБ"] = df["ТБ"].astype("string").fillna("").str.strip()
        df["ГОСБ"] = df["ГОСБ"].astype("string").fillna("").str.strip()
        df["КМ"] = df["КМ"].astype("string").fillna("").str.strip()

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

        elapsed = time.perf_counter() - start_time_perf
        file_stat["Время загрузки (сек)"] = elapsed
        log_debug(f"Файл загружен успешно: {file_name}", def_name="load_single_file")
        return df, file_stat
    except Exception as exc:
        file_stat["Статус"] = "ОШИБКА"
        file_stat["Ошибка"] = str(exc)
        file_stat["Время загрузки (сек)"] = time.perf_counter() - start_time_perf
        LOGGER.error("Ошибка загрузки файла %s: %s", file_name, exc)
        log_debug(f"Ошибка загрузки файла {file_name}: {exc}", def_name="load_single_file")
        return None, file_stat


def load_orgunit_mapping() -> pd.DataFrame:
    """Загружает первый CSV из IN/OrgUnit и возвращает маппинг ГОСБ -> Кластер."""
    if not os.path.exists(ORGUNIT_FOLDER):
        LOGGER.warning("Папка OrgUnit не найдена: %s", ORGUNIT_FOLDER)
        return pd.DataFrame(columns=["Код ГОСБ", "Кластер"])

    csv_files = glob.glob(os.path.join(ORGUNIT_FOLDER, ORGUNIT_FILE_GLOB))
    if not csv_files:
        LOGGER.warning("CSV-файлы в OrgUnit не найдены: %s", ORGUNIT_FOLDER)
        return pd.DataFrame(columns=["Код ГОСБ", "Кластер"])

    # Файлы с подстрокой TEST в имени идут в конец — чтобы не перекрывать полный справочник.
    csv_files = sorted(
        csv_files,
        key=lambda p: ("TEST" in os.path.basename(p).upper(), os.path.basename(p).lower()),
    )
    org_path = csv_files[0]
    LOGGER.info("Загрузка OrgUnit: %s", os.path.basename(org_path))
    log_debug(f"Читаем OrgUnit файл: {org_path}", def_name="load_orgunit_mapping")

    # Пытаемся прочитать как ';', если не подходит - читаем как ','.
    org_df = pd.read_csv(org_path, sep=";", dtype=str, encoding="utf-8", engine="python")
    org_df.columns = [str(c).strip().lstrip("\ufeff") for c in org_df.columns]
    if "GOSB_CODE" not in org_df.columns:
        org_df = pd.read_csv(org_path, sep=",", dtype=str, encoding="utf-8", engine="python")
        org_df.columns = [str(c).strip().lstrip("\ufeff") for c in org_df.columns]

    cluster_src: str | None = None
    if "CLUSTER" in org_df.columns:
        cluster_src = "CLUSTER"
    elif "CLUSTER_CODE" in org_df.columns:
        cluster_src = "CLUSTER_CODE"

    if "GOSB_CODE" not in org_df.columns or cluster_src is None:
        raise ValueError(
            "Файл OrgUnit должен содержать колонку GOSB_CODE и одну из: CLUSTER, CLUSTER_CODE."
        )

    org_df = org_df[["GOSB_CODE", cluster_src]].copy()
    org_df.rename(columns={cluster_src: "CLUSTER"}, inplace=True)
    org_df["Код ГОСБ"] = normalize_gosb_code(org_df["GOSB_CODE"])
    org_df["Кластер"] = org_df["CLUSTER"].astype("string").fillna("").str.strip()
    org_df = org_df[["Код ГОСБ", "Кластер"]].drop_duplicates()
    LOGGER.info(
        "OrgUnit: уникальных кодов ГОСБ в справочнике: %d",
        org_df["Код ГОСБ"].nunique(),
    )
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
    total_start_perf = time.perf_counter()
    resolved_workers, workers_mode = resolve_max_workers(MAX_WORKERS_CONFIG)
    LOGGER.info("Старт обработки ОД")
    log_debug("Старт обработки ОД", def_name="main")

    cprint("=" * 70)
    cprint("НАЧАЛО ОБРАБОТКИ (ВЕРСИЯ v3.1)")
    cprint("=" * 70)
    cprint(f"Время запуска:         {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    cprint(f"Рабочая папка:         {SCRIPT_DIR}", level="verbose")
    cprint(f"Входная папка (IN):    {INPUT_FOLDER}", level="verbose")
    cprint(f"Папка OrgUnit:         {ORGUNIT_FOLDER}", level="verbose")
    cprint(f"Выходная папка (OUT):  {RUN_OUTPUT_FOLDER}")
    cprint(f"Папка логов:           {RUN_LOG_FOLDER}")
    cprint(f"Таймштамп:             {TIMESTAMP}", level="verbose")
    cprint(f"Потоков загрузки:      {resolved_workers} (режим: {workers_mode})")
    cprint(f"Уровень консоли:       {CONSOLE_VERBOSITY}")

    if not os.path.exists(INPUT_FOLDER):
        raise FileNotFoundError("Папка IN не существует.")

    file_list = sorted(glob.glob(os.path.join(INPUT_FOLDER, INPUT_FILE_GLOB)))
    file_list = [f for f in file_list if not os.path.basename(f).startswith("~$")]
    if not file_list:
        raise FileNotFoundError("В папке IN нет XLSX-файлов.")

    cprint(f"\nНайдено XLSX файлов: {len(file_list)}")
    for idx, fp in enumerate(file_list, start=1):
        cprint(f"  [{idx:2d}] {os.path.basename(fp)}", level="verbose")

    dfs: list[pd.DataFrame] = []
    stats: list[dict[str, Any]] = []

    cprint("\n" + "=" * 70)
    cprint(f"ЗАГРУЗКА ФАЙЛОВ (параллельно, {resolved_workers} потоков)")
    cprint("=" * 70)

    stage_start = time.perf_counter()
    with ThreadPoolExecutor(max_workers=resolved_workers) as executor:
        future_to_file = {executor.submit(load_single_file, path): path for path in file_list}
        completed = 0
        for future in as_completed(future_to_file):
            file_name = os.path.basename(future_to_file[future])
            completed += 1
            df_part, file_stat = future.result()
            stats.append(file_stat)
            if df_part is not None:
                dfs.append(df_part)
                elapsed_total = time.perf_counter() - total_start_perf
                cprint(
                    f"[{completed:2d}/{len(file_list)}] ✓ {file_name[:45]:<45} | "
                    f"исх: {file_stat['Строк исходных']:>8,} | "
                    f"удалено(СЗ): {file_stat['Удалено серая зона']:>7,} | "
                    f"осталось: {file_stat['Строк после фильтра']:>8,} | "
                    f"файл: {fmt_elapsed(file_stat['Время загрузки (сек)'])} | "
                    f"с начала: {fmt_elapsed(elapsed_total)}",
                    level="normal",
                )
            else:
                cprint(f"[{completed:2d}/{len(file_list)}] ✗ {file_name}: {file_stat['Ошибка'][:55]}")
    stage_elapsed = time.perf_counter() - stage_start
    total_elapsed = time.perf_counter() - total_start_perf
    cprint(f"⏱ Этап загрузки: {fmt_elapsed(stage_elapsed)} | с начала: {fmt_elapsed(total_elapsed)}")

    if not dfs:
        raise RuntimeError("Нет успешно загруженных файлов.")

    cprint("\nОбъединение данных...")
    stage_start = time.perf_counter()
    df_combined = pd.concat(dfs, ignore_index=True, copy=False)
    del dfs
    cprint(f"Объединено строк: {len(df_combined):,}")
    stage_elapsed = time.perf_counter() - stage_start
    total_elapsed = time.perf_counter() - total_start_perf
    cprint(f"⏱ Этап объединения: {fmt_elapsed(stage_elapsed)} | с начала: {fmt_elapsed(total_elapsed)}")

    total_grey_removed = sum(s["Удалено серая зона"] for s in stats if s["Статус"] == "OK")
    errors_date = int(df_combined["Дата загрузки"].isna().sum())
    errors_prev = int(df_combined["ПРОШЛЫЙ ГОД ОД, тыс. руб."].isna().sum())
    errors_curr = int(df_combined["ТЕКУЩИЙ ГОД ОД, тыс. руб."].isna().sum())
    total_errors = errors_date + errors_prev + errors_curr

    output_raw_base = os.path.join(RUN_OUTPUT_FOLDER, f"{RAW_BASENAME}_{TIMESTAMP}")
    output_agg_base = os.path.join(RUN_OUTPUT_FOLDER, f"{AGG_BASENAME}_{TIMESTAMP}")
    output_last_km_base = os.path.join(RUN_OUTPUT_FOLDER, f"{LAST_KM_BASENAME}_{TIMESTAMP}")
    output_dyn_base = os.path.join(RUN_OUTPUT_FOLDER, f"{DYN_BASENAME}_{TIMESTAMP}")
    output_final_base = os.path.join(RUN_OUTPUT_FOLDER, f"{FINAL_BASENAME}_{TIMESTAMP}")
    output_stats_base = os.path.join(RUN_OUTPUT_FOLDER, f"{STATS_BASENAME}_{TIMESTAMP}")
    output_all_xlsx = os.path.join(RUN_OUTPUT_FOLDER, f"00_all_results_{TIMESTAMP}.xlsx")

    cprint("\nСохранение сырого слоя...", level="verbose")
    stage_start = time.perf_counter()
    cprint(f"raw-слой подготовлен, строк: {len(df_combined):,}", level="verbose")
    stage_elapsed = time.perf_counter() - stage_start
    total_elapsed = time.perf_counter() - total_start_perf
    cprint(
        f"⏱ Этап сохранения raw: {fmt_elapsed(stage_elapsed)} | с начала: {fmt_elapsed(total_elapsed)}",
        level="verbose",
    )

    cprint("\nАгрегация по ключу...")
    stage_start = time.perf_counter()
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
    cprint(f"Агрегированных строк: {len(agg_df):,}")
    cprint("Агрегат подготовлен", level="verbose")
    stage_elapsed = time.perf_counter() - stage_start
    total_elapsed = time.perf_counter() - total_start_perf
    cprint(f"⏱ Этап агрегации: {fmt_elapsed(stage_elapsed)} | с начала: {fmt_elapsed(total_elapsed)}")

    cprint("\nОпределение последнего КМ...")
    stage_start = time.perf_counter()
    df_for_last_km = df_combined[df_combined["Дата загрузки"].notna()].copy()
    max_keys = ["ИНН_12", "Код ТБ", "Код ГОСБ", "Дата загрузки"]
    curr_fill = df_for_last_km["ТЕКУЩИЙ ГОД ОД, тыс. руб."].fillna(float("-inf"))
    idx_max = curr_fill.groupby([df_for_last_km[k] for k in max_keys], sort=False, dropna=False).idxmax()
    grouped_by_date = df_for_last_km.loc[idx_max.values].reset_index(drop=True)

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
    cprint(f"Строк в 03_last_km_by_inn_tb: {len(last_km_with_dates):,}")
    stage_elapsed = time.perf_counter() - stage_start
    total_elapsed = time.perf_counter() - total_start_perf
    cprint(f"⏱ Этап последнего КМ: {fmt_elapsed(stage_elapsed)} | с начала: {fmt_elapsed(total_elapsed)}")

    cprint("\nРасчет динамики по последнему КМ...")
    stage_start = time.perf_counter()
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
    prev_vals = dyn_group["Сумма ПРОШЛЫЙ ГОД ОД, тыс. руб."].to_numpy()
    curr_vals = dyn_group["Сумма ТЕКУЩИЙ ГОД ОД, тыс. руб."].to_numpy()
    nan_mask = np.isnan(prev_vals) | np.isnan(curr_vals)
    zero_prev_mask = prev_vals == 0
    base_rate = (curr_vals - prev_vals) / np.abs(prev_vals) * 100.0
    zero_rate = np.where(curr_vals > 0, 100.0, np.where(curr_vals < 0, -100.0, 0.0))
    growth_rate = np.where(zero_prev_mask, zero_rate, base_rate)
    growth_rate = np.where(nan_mask, np.nan, growth_rate)
    dyn_group["Темп прироста, %"] = growth_rate

    cprint(f"Строк в 04_km_dynamics: {len(dyn_group):,}")
    stage_elapsed = time.perf_counter() - stage_start
    total_elapsed = time.perf_counter() - total_start_perf
    cprint(f"⏱ Этап динамики: {fmt_elapsed(stage_elapsed)} | с начала: {fmt_elapsed(total_elapsed)}")

    cprint("\nЗагрузка OrgUnit и обогащение кластером...")
    stage_start = time.perf_counter()
    orgunit_map = load_orgunit_mapping()
    final_result = pd.merge(dyn_group, orgunit_map, on="Код ГОСБ", how="left")
    final_result["Кластер"] = final_result["Кластер"].fillna("НЕ НАЙДЕН")
    # Медианы по всем строкам с тем же кластером (для сравнения строки с «типичным» уровнем кластера).
    # В именах сохраняем «Прирост» / «Темп прироста», чтобы форматирование листа совпало с остальными колонками.
    col_med_growth = "Прирост ОД по кластеру (медиана), тыс. руб."
    col_med_rate = "Темп прироста по кластеру (медиана), %"
    final_result[col_med_growth] = final_result.groupby("Кластер", dropna=False)[
        "Прирост ОД, тыс. руб."
    ].transform("median")
    final_result[col_med_rate] = final_result.groupby("Кластер", dropna=False)[
        "Темп прироста, %"
    ].transform("median")
    cols_base = [c for c in final_result.columns if c not in (col_med_growth, col_med_rate)]
    pos = cols_base.index("Кластер") + 1
    final_result = final_result[cols_base[:pos] + [col_med_growth, col_med_rate] + cols_base[pos:]]
    matched_cluster = int((final_result["Кластер"] != "НЕ НАЙДЕН").sum())
    LOGGER.info(
        "Кластер: совпало со справочником OrgUnit %d строк из %d",
        matched_cluster,
        len(final_result),
    )
    cprint(f"Строк в 05_final_result_with_cluster: {len(final_result):,}")
    cprint(f"Непросопоставленных кластеров: {int((final_result['Кластер'] == 'НЕ НАЙДЕН').sum()):,}")
    stage_elapsed = time.perf_counter() - stage_start
    total_elapsed = time.perf_counter() - total_start_perf
    cprint(f"⏱ Этап кластера: {fmt_elapsed(stage_elapsed)} | с начала: {fmt_elapsed(total_elapsed)}")

    # ===== ФОРМИРОВАНИЕ СТАТИСТИКИ + ЭКСПОРТ =====
    cprint("\nФормирование статистики и итогового экспорта...")
    stage_start = time.perf_counter()
    stats.sort(key=lambda x: x["Файл"])
    total_minutes = round((datetime.now() - total_start).total_seconds() / 60, 2)
    stats_files_df, stats_summary_df = build_stats_frames(
        stats=stats,
        timestamp=TIMESTAMP,
        workers=resolved_workers,
        workers_mode=workers_mode,
        df_combined=df_combined,
        total_grey_removed=total_grey_removed,
        agg_df=agg_df,
        last_km_result=last_km_result,
        last_km_with_dates=last_km_with_dates,
        dyn_group=dyn_group,
        final_result=final_result,
        total_errors=total_errors,
        orgunit_map=orgunit_map,
        total_minutes=total_minutes,
    )

    export_items: list[tuple[str, pd.DataFrame, str]] = [
        ("01_raw_combined", df_combined, output_raw_base),
        ("02_aggregated", agg_df, output_agg_base),
        ("03_last_km", last_km_with_dates, output_last_km_base),
        ("04_km_dynamics", dyn_group, output_dyn_base),
        ("05_final_cluster", final_result, output_final_base),
        ("06_stats_files", stats_files_df, f"{output_stats_base}_files"),
        ("07_stats_summary", stats_summary_df, f"{output_stats_base}_summary"),
    ]

    xlsx_items = [(sheet, frame) for sheet, frame, _ in export_items if should_export_to_xlsx(len(frame))]
    csv_items = [(sheet, frame, base) for sheet, frame, base in export_items if not should_export_to_xlsx(len(frame))]

    cprint(
        f"Режим экспорта: смешанный | XLSX-листов: {len(xlsx_items)} | CSV-файлов: {len(csv_items)}"
    )

    if xlsx_items:
        wb = Workbook()
        wb.remove(wb.active)
        for sheet, frame in xlsx_items:
            write_df_to_sheet(wb, sheet, frame)
        wb.save(output_all_xlsx)
        cprint(f"✓ {os.path.basename(output_all_xlsx)}")
    else:
        cprint("XLSX не создан: все таблицы превышают лимит строк.")

    for sheet, frame, base in csv_items:
        frame.to_csv(f"{base}.csv", index=False, sep=";", decimal=",", encoding="utf-8-sig")
        cprint(f"✓ {os.path.basename(base)}.csv ({sheet})", level="verbose")

    stage_elapsed = time.perf_counter() - stage_start
    total_elapsed = time.perf_counter() - total_start_perf
    cprint(f"⏱ Этап экспорта: {fmt_elapsed(stage_elapsed)} | с начала: {fmt_elapsed(total_elapsed)}")

    total_time = (datetime.now() - total_start).total_seconds()
    LOGGER.info("Обработка завершена успешно за %.2f сек", total_time)
    log_debug("Обработка завершена успешно", def_name="main")

    cprint("\n" + "=" * 70)
    cprint("ОБРАБОТКА ЗАВЕРШЕНА")
    cprint("=" * 70)
    cprint(f"⏱ ОБЩЕЕ ВРЕМЯ: {total_time:.1f} сек ({total_time / 60:.2f} мин)")
    cprint("\nВыходные файлы в OUT:")
    if xlsx_items:
        cprint(f"  XLSX: {os.path.basename(output_all_xlsx)}")
    else:
        cprint("  XLSX: не создан")
    if csv_items:
        cprint("  CSV:")
        for _, _, base in csv_items:
            cprint(f"    - {os.path.basename(base)}.csv", level="normal")
    else:
        cprint("  CSV: не созданы")
    cprint("=" * 70)


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        LOGGER.exception("Критическая ошибка выполнения: %s", exc)
        print(f"\n✗ КРИТИЧЕСКАЯ ОШИБКА: {exc}")
        raise