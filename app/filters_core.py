import pandas as pd
from dataclasses import dataclass
from enum import Enum


class Operator(str, Enum):
    CONTAINS = "contains"
    EQUALS = "equals"
    NOT_EQUALS = "not_equals"
    GREATER = "greater"
    LESS = "less"
    RANGE = "range"
    IS_TRUE = "is_true"
    IS_FALSE = "is_false"
    NOT_NULL = "not_null"
    IS_NULL = "is_null"


@dataclass
class FilterCondition:
    column: str
    operator: Operator
    value: any = None


def apply_filters(df: pd.DataFrame, conditions: list[FilterCondition]) -> pd.DataFrame:
    if not conditions:
        return df

    mask = pd.Series([True] * len(df), index=df.index)

    for cond in conditions:
        col = df[cond.column]

        if cond.operator == Operator.CONTAINS:
            m = col.astype(str).str.contains(str(cond.value), case=False, na=False)

        elif cond.operator == Operator.EQUALS:
            m = col == cond.value

        elif cond.operator == Operator.NOT_EQUALS:
            m = col != cond.value

        elif cond.operator == Operator.GREATER:
            m = col > cond.value

        elif cond.operator == Operator.LESS:
            m = col < cond.value

        elif cond.operator == Operator.RANGE:
            start, end = cond.value
            m = pd.Series([True] * len(df), index=df.index)
            if start is not None:
                m &= col >= start
            if end is not None:
                m &= col <= end

        elif cond.operator == Operator.IS_TRUE:
            m = col == True

        elif cond.operator == Operator.IS_FALSE:
            m = col == False

        elif cond.operator == Operator.NOT_NULL:
            m = col.notna()

        elif cond.operator == Operator.IS_NULL:
            m = col.isna()

        else:
            m = pd.Series([True] * len(df), index=df.index)

        mask &= m

    return df[mask]
