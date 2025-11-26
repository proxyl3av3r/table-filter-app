import os
import pandas as pd
from docx import Document


def _load_from_docx(path: str) -> pd.DataFrame:
    """
    Загрузить первую осмысленную таблицу из .docx в DataFrame.
    Ожидается прямоугольная таблица (как твой реєстр).
    """
    doc = Document(path)

    table = None
    for t in doc.tables:
        if len(t.rows) >= 2:
            table = t
            break

    if table is None:
        raise ValueError("У документі Word не знайдено підходящої таблиці.")

    headers = [cell.text.strip() for cell in table.rows[0].cells]
    rows = []

    for row in table.rows[1:]:
        values = [cell.text.strip() for cell in row.cells]

        if len(values) < len(headers):
            values += [""] * (len(headers) - len(values))
        elif len(values) > len(headers):
            values = values[: len(headers)]

        rows.append(dict(zip(headers, values)))

    df = pd.DataFrame(rows)
    return df


def load_test_df(path: str = "registry_test.csv") -> pd.DataFrame:
    """
    Универсальная загрузка:
    - .csv
    - .xlsx / .xls
    - .docx (таблица Word)
    """
    ext = os.path.splitext(path)[1].lower()

    if ext == ".csv":
        df = pd.read_csv(path)
    elif ext in (".xlsx", ".xls"):
        df = pd.read_excel(path)
    elif ext == ".docx":
        df = _load_from_docx(path)
    else:
        df = pd.read_csv(path)

    # Примерный набор дат (под твой реєстр, можно расширить по факту)
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

    # Булеві (Так/Ні) – можно расширить список
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