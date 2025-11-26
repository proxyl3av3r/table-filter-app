import datetime
from typing import Any, Callable, Optional, Set

import numpy as np
import pandas as pd
from PySide6.QtCore import QAbstractTableModel, QModelIndex, Qt
from PySide6.QtGui import QBrush, QColor


class PandasTableModel(QAbstractTableModel):
    """
    Табличная модель для отображения pandas.DataFrame в QTableView.

    Возможности:
    - отображение и редактирование данных;
    - auto-format дат -> дд.мм.рррр;
    - отображение boolean -> Так / Ні;
    - подсветка строк, где сроки спливают (highlight_indices).
    - обратная связь через edit_callback — изменение отражается в df_original.
    """

    def __init__(
        self,
        df: pd.DataFrame,
        parent=None,
        edit_callback: Optional[Callable[[Any, str, Any], None]] = None,
        highlight_indices: Optional[Set[Any]] = None,
    ):
        super().__init__(parent)
        self._df = df
        self._edit_callback = edit_callback
        self._highlight_indices: Set[Any] = highlight_indices or set()

    # =============================
    # Базовые методы модели
    # =============================

    def rowCount(self, parent=QModelIndex()) -> int:
        if parent.isValid():
            return 0
        return len(self._df)

    def columnCount(self, parent=QModelIndex()) -> int:
        if parent.isValid():
            return 0
        return len(self._df.columns)

    # =============================
    # Отображение данных
    # =============================

    def data(self, index: QModelIndex, role=Qt.DisplayRole) -> Any:
        if not index.isValid():
            return None

        row = index.row()
        col = index.column()
        series = self._df.iloc[:, col]

        if role in (Qt.DisplayRole, Qt.EditRole):
            value = self._df.iat[row, col]

            if pd.isna(value):
                return ""

            # даты -> дд.мм.рррр
            if isinstance(value, (pd.Timestamp, datetime.datetime, datetime.date)):
                try:
                    return value.strftime("%d.%m.%Y")
                except Exception:
                    return str(value)

            # bool -> Так / Ні
            if isinstance(value, (bool, np.bool_)):
                return "Так" if value else "Ні"

            return str(value)

        # =============================
        # Подсветка строк (expiring)
        # =============================
        if role == Qt.BackgroundRole and self._highlight_indices:
            try:
                orig_index = self._df.index[row]
                if orig_index in self._highlight_indices:
                    # Полупрозрачный красный
                    return QBrush(QColor(255, 0, 0, 60))
            except Exception:
                pass

        return None

    # =============================
    # Заголовки
    # =============================

    def headerData(self, section: int, orientation, role=Qt.DisplayRole):
        if role != Qt.DisplayRole:
            return None

        if orientation == Qt.Horizontal:
            return str(self._df.columns[section])
        else:
            return str(section + 1)

    # =============================
    # Редактирование
    # =============================

    def flags(self, index: QModelIndex):
        if not index.isValid():
            return Qt.ItemIsEnabled
        return Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsEditable

    def setData(self, index: QModelIndex, value, role=Qt.EditRole) -> bool:
        if role != Qt.EditRole or not index.isValid():
            return False

        row = index.row()
        col = index.column()
        col_name = self._df.columns[col]
        series = self._df[col_name]

        text = str(value).strip()
        new_val: Any = text

        try:
            # datetime
            if pd.api.types.is_datetime64_any_dtype(series):
                if text == "":
                    new_val = pd.NaT
                else:
                    new_val = pd.to_datetime(text, format="%d.%m.%Y", dayfirst=True)

            # boolean
            elif pd.api.types.is_bool_dtype(series):
                if text == "":
                    new_val = False
                else:
                    t = text.lower()
                    new_val = t in ("так", "true", "1", "yes", "y", "+")

            # numeric
            elif pd.api.types.is_numeric_dtype(series):
                if text == "":
                    new_val = np.nan
                else:
                    try:
                        new_val = int(text)
                    except ValueError:
                        new_val = float(text)

            # everything else
            else:
                new_val = text

        except Exception:
            new_val = text

        # Записываем в DataFrame
        self._df.iat[row, col] = new_val

        # callback — синхронизация с df_original в MainWindow
        if self._edit_callback is not None:
            orig_index = self._df.index[row]
            self._edit_callback(orig_index, col_name, new_val)

        # уведомляем View об изменении
        self.dataChanged.emit(index, index, [Qt.DisplayRole, Qt.EditRole])
        return True

    # =============================
    # Полная замена DataFrame
    # =============================

    def update_df(self, df: pd.DataFrame, highlight_indices: Optional[Set[Any]] = None):
        """
        Обновляет внутренний df модели + список строк, которые подсвечиваются.
        """
        self.beginResetModel()
        self._df = df
        if highlight_indices is not None:
            self._highlight_indices = highlight_indices
        self.endResetModel()

    @property
    def df(self) -> pd.DataFrame:
        return self._df