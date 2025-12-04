import os
import pandas as pd
from docx import Document


def _load_from_docx(path: str) -> pd.DataFrame:
    """
    Загрузить все однотипные таблицы из .docx в один DataFrame.

    Логика:
    - Берём первую «осмислену» таблицу как эталонную (там заголовки).
    - По всем остальным таблицам:
        * проверяем совпадение количества колонок;
        * пропускаем их первую строку (свою шапку);
        * добавляем строки в общий список.
    - Полностью пустые строки отбрасываем.
    """
    doc = Document(path)

    header: list[str] | None = None
    rows: list[dict[str, str]] = []

    for table in doc.tables:
        # Таблица должна хотя бы иметь шапку + одну строку данных
        if len(table.rows) < 2:
            continue

        # Шапка текущей таблицы
        current_header = [cell.text.strip() for cell in table.rows[0].cells]

        # Если шапка полностью пустая — пропускаем таблицу
        if not any(current_header):
            continue

        # Первая «нормальная» таблица — задаём эталонный header
        if header is None:
            header = current_header
        else:
            # Если количество колонок не совпадает — считаем,
            # что это другая структура, такую таблицу пропускаем.
            if len(current_header) != len(header):
                continue

        # Обрабатываем строки данных (пропускаем собственный header таблицы)
        for row in table.rows[1:]:
            values = [cell.text.strip() for cell in row.cells]

            # Выравниваем длину под header
            if len(values) < len(header):
                values += [""] * (len(header) - len(values))
            elif len(values) > len(header):
                values = values[: len(header)]

            # Если строка полностью пустая — пропускаем
            if not any(values):
                continue

            rows.append(dict(zip(header, values)))

    if header is None:
        raise ValueError("У документі Word не знайдено придатних таблиць.")

    df = pd.DataFrame(rows)
    return df


def load_test_df(path: str = "registry_test.csv") -> pd.DataFrame:
    """
    Универсальная загрузка:
    - .csv
    - .xlsx / .xls
    - .docx (одна или несколько однотипных таблиц Word в один DataFrame)
    """
    ext = os.path.splitext(path)[1].lower()

    if ext == ".csv":
        df = pd.read_csv(path)
    elif ext in (".xlsx", ".xls"):
        df = pd.read_excel(path)
    elif ext == ".docx":
        df = _load_from_docx(path)
    else:
        # по умолчанию пытаемся как csv
        df = pd.read_csv(path)

    # Примерный набор колонок-дат (под твой реєстр, можно дополнять)
    date_cols = [
        "Дата_реєстрації",
        "Дата_нар",
        "Дата_повідомлення_підозри",
        "Дата_зупинення",
        "Дата_доручення_розшуку",
        "Дата_заведення_ОРС",
        "Дата_інфо_про_перетин",
        "Дата_адмін",
        "Дата_міжнар_розшуку",
        "Дата_виїзду",
        "Дата_оголошення_в_розшук",
    ]
    for col in date_cols:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce", dayfirst=True)

    # Булеві (Так/Ні) – можна доповнювати список при потребі
    bool_map = {
        "Так": True,
        "так": True,
        "ТАК": True,
        "Ні": False,
        "ні": False,
        "НІ": False,
        "True": True,
        "False": False,
        True: True,
        False: False,
    }
    bool_cols = [
        "Є_виїзд_за_кордон",
        "Є_Інтерпол",
        "Є_інфо_про_перетин_кордону",
        "Є_адмін_відповідальність",
    ]
    for bcol in bool_cols:
        if bcol in df.columns:
            df[bcol] = df[bcol].map(bool_map).fillna(False).astype(bool)

    return df