import os
import pandas as pd
from docx import Document


def _load_from_docx(path: str) -> pd.DataFrame:
    """
    Загрузить первую "нормальную" таблицу из .docx в DataFrame.
    Ожидается прямоугольная таблица без хитрых объединений (как мы сами экспортируем из приложения).
    """
    doc = Document(path)

    # берём первую таблицу, где есть хотя бы 2 строки (заголовок + данные)
    table = None
    for t in doc.tables:
        if len(t.rows) >= 2:
            table = t
            break

    if table is None:
        raise ValueError("У документі Word не знайдено підходящої таблиці.")

    # заголовки — первая строка
    headers = [cell.text.strip() for cell in table.rows[0].cells]

    rows = []
    for row in table.rows[1:]:
        values = [cell.text.strip() for cell in row.cells]

        # подгоняем длину под количество колонок
        if len(values) < len(headers):
            values += [""] * (len(headers) - len(values))
        elif len(values) > len(headers):
            values = values[: len(headers)]

        rows.append(dict(zip(headers, values)))

    df = pd.DataFrame(rows)
    return df


def load_test_df(path: str = "registry_test.csv") -> pd.DataFrame:
    """
    Универсальный загрузчик:
    - .csv
    - .xlsx / .xls
    - .docx (Word-таблица)
    """
    ext = os.path.splitext(path)[1].lower()

    if ext == ".csv":
        df = pd.read_csv(path)
    elif ext in (".xlsx", ".xls"):
        df = pd.read_excel(path)
    elif ext == ".docx":
        df = _load_from_docx(path)
    else:
        # по умолчанию пробуем csv
        df = pd.read_csv(path)

    # Приводим даты к datetime (чтобы фильтры по датам работали)
    date_cols = ["Дата_нар", "Дата_виїзду", "Дата_оголошення_в_розшук"]
    for col in date_cols:
        if col in df.columns:
            # умеет парсить и форматы типу "2024-01-01", и "01.01.2024"
            df[col] = pd.to_datetime(df[col], errors="coerce", dayfirst=True)

    # Булеві поля (наш експорт пише "Так"/"Ні")
    bool_map = {"Так": True, "Ні": False, "True": True, "False": False}
    for bcol in ["Є_виїзд_за_кордон", "Є_Інтерпол"]:
        if bcol in df.columns:
            df[bcol] = df[bcol].map(bool_map).fillna(False).astype(bool)

    return df