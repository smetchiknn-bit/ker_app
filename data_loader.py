"""
Модуль загрузки и обработки данных КЕР.
Воспроизводит логику VBA-макроса "ВыборкаИзВыгрузкиСтикера_Упрощенная".
"""
import pandas as pd
from pathlib import Path


# Столбцы, которые мы ищем в исходном листе "КЕР в ССР"
RAW_COLUMNS = [
    "ИД_КЕР", "Наименование_КЕР", "ЕдИзм КЕР",
    "Лист состава работ КЕР",
    "Л1 Код", "Л1 Наименование",
    "Л2 Код", "Л2 Наименование",
    "Л3 Код", "Л3 Наименование",
    "Л4 Код", "Л4 Наименование",
    "Л5 Код", "Л5 Наименование",
    "ФЕР"
]

# Итоговый порядок столбцов на листе "Выгрузка"
TARGET_COLUMNS = [
    "ИД_КЕР", "Наименование_КЕР", "ЕдИзм КЕР", "Иерархия",
    "Лист состава работ КЕР",
    "Л1 Код", "Л1 Наименование",
    "Л2 Код", "Л2 Наименование",
    "Л3 Код", "Л3 Наименование",
    "Л4 Код", "Л4 Наименование",
    "Л5 Код", "Л5 Наименование",
    "ФЕР"
]


def _build_hierarchy(row: pd.Series) -> str:
    """
    Формирует текстовую иерархию из кодов и наименований Л1-Л5.
    Аналог VBA-цикла в макросе.
    """
    lines = []
    for level in range(1, 6):
        code = row.get(f"Л{level} Код", "")
        name = row.get(f"Л{level} Наименование", "")
        code = "" if pd.isna(code) else str(code).strip()
        name = "" if pd.isna(name) else str(name).strip()
        if code or name:
            lines.append(f"{code}. {name}".strip(". "))
    return "\n".join(lines)


def process_raw_sheet(df_raw: pd.DataFrame) -> pd.DataFrame:
    """
    Применяет к "сырому" DataFrame (лист 'КЕР в ССР') логику макроса:
    1. Отбирает нужные столбцы.
    2. Формирует столбец 'Иерархия'.
    """
    # Берём только те колонки, которые реально есть в файле
    available = [c for c in RAW_COLUMNS if c in df_raw.columns]
    missing = [c for c in RAW_COLUMNS if c not in df_raw.columns]
    if missing:
        print(f"[WARN] Отсутствуют колонки в источнике: {missing}")

    df = df_raw[available].copy()

    # Формируем Иерархию
    df["Иерархия"] = df.apply(_build_hierarchy, axis=1)

    # Приводим к итоговому порядку (пропуская отсутствующие)
    final_cols = [c for c in TARGET_COLUMNS if c in df.columns]
    df = df[final_cols]

    # Приводим ИД_КЕР к строке (чтобы не было проблем с float)
    if "ИД_КЕР" in df.columns:
        df["ИД_КЕР"] = df["ИД_КЕР"].astype(str).str.replace(".0", "", regex=False)

    return df


def load_or_process(file_path: str, target_sheet: str = "Выгрузка") -> tuple[pd.DataFrame, str]:
    """
    Загружает данные из Excel.
    - Если есть лист target_sheet ("Выгрузка") — берём его как есть.
    - Иначе ищем лист "КЕР в ССР" и применяем обработку.
    Возвращает (DataFrame, source_sheet_name).
    """
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"Файл не найден: {file_path}")

    xls = pd.ExcelFile(file_path, engine="openpyxl")
    sheet_names = xls.sheet_names

    if target_sheet in sheet_names:
        df = pd.read_excel(xls, sheet_name=target_sheet)
        return df, target_sheet

    # Ищем "КЕР в ССР" (с учётом возможных опечаток)
    raw_sheet = None
    for name in sheet_names:
        if "КЕР в ССР" in name or "КЕР в СР" in name:
            raw_sheet = name
            break

    if raw_sheet is None:
        raise ValueError(
            f"В файле нет ни листа '{target_sheet}', ни листа 'КЕР в ССР'.\n"
            f"Доступные листы: {sheet_names}"
        )

    df_raw = pd.read_excel(xls, sheet_name=raw_sheet)
    df = process_raw_sheet(df_raw)
    return df, raw_sheet


def export_processed(df: pd.DataFrame, output_path: str) -> None:
    """Экспортирует обработанную таблицу в Excel с форматированием."""
    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name="Выгрузка", index=False)

        workbook = writer.book
        worksheet = writer.sheets["Выгрузка"]

        # Формат заголовка
        header_fmt = workbook.add_format({
            "bold": True, "align": "center", "valign": "vcenter",
            "border": 1, "bg_color": "#D9E1F2"
        })
        # Формат данных
        cell_fmt = workbook.add_format({
            "valign": "top", "text_wrap": True, "border": 1
        })

        # Применяем форматы
        for col_num, col_name in enumerate(df.columns):
            worksheet.write(0, col_num, col_name, header_fmt)
            worksheet.set_column(col_num, col_num, 20, cell_fmt)

        # Закрепляем первую строку
        worksheet.freeze_panes(1, 0)


def find_duplicate_ids(df: pd.DataFrame) -> pd.DataFrame:
    """
    Находит ИД_КЕР, у которых несколько разных Наименований/ЕдИзм/Иерархий.
    Возвращает DataFrame с дублями.
    """
    if "ИД_КЕР" not in df.columns:
        return pd.DataFrame()

    grouped = df.groupby("ИД_КЕР").agg({
        "Наименование_КЕР": "nunique",
        "ЕдИзм КЕР": "nunique",
        "Иерархия": "nunique"
    }).reset_index()

    dupes = grouped[
        (grouped["Наименование_КЕР"] > 1) |
        (grouped["ЕдИзм КЕР"] > 1) |
        (grouped["Иерархия"] > 1)
    ]
    return dupes
