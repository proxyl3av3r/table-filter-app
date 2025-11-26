from dataclasses import dataclass
from enum import Enum, auto
from typing import Any, List, Tuple

import pandas as pd


class Operator(Enum):
    CONTAINS = auto()
    EQUALS = auto()
    NOT_EQUALS = auto()
    RANGE = auto()  # в основному для діапазону дат


@dataclass
class FilterCondition:
    column: str
    operator: Operator
    value: Any  # для RANGE очікуємо Tuple[start, end]


def _apply_single_condition(df: pd.DataFrame, cond: FilterCondition) -> pd.DataFrame:
    col = cond.column
    if col not in df.columns:
        return df

    series = df[col]
    op = cond.operator
    val = cond.value

    # --------- Діапазон (включно з датами в текстових полях) ---------
    if op == Operator.RANGE:
        start, end = val  # (можуть бути None / None)

        # Якщо стовпець уже datetime64 – використовуємо напряму
        if pd.api.types.is_datetime64_any_dtype(series):
            date_series = series
        else:
            # Спроба витягнути першу дату формату дд.мм.рррр з тексту
            extracted = series.astype(str).str.extract(r"(\d{2}\.\d{2}\.\d{4})")[0]
            date_series = pd.to_datetime(
                extracted, format="%d.%m.%Y", errors="coerce"
            )

        mask = pd.Series(True, index=df.index)

        if start is not None:
            mask &= date_series >= start
        if end is not None:
            mask &= date_series <= end

        return df[mask]

    # --------- Містить ---------
    if op == Operator.CONTAINS:
        text = str(val)
        return df[series.astype(str).str.contains(text, case=False, na=False)]

    # --------- Дорівнює ---------
    if op == Operator.EQUALS:
        return df[series == val]

    # --------- Не дорівнює ---------
    if op == Operator.NOT_EQUALS:
        return df[series != val]

    return df


def apply_filters(df: pd.DataFrame, conditions: List[FilterCondition]) -> pd.DataFrame:
    """Застосувати список умов послідовно."""
    result = df
    for cond in conditions:
        result = _apply_single_condition(result, cond)
        if result.empty:
            break
    return result