import datetime
from typing import Any, Callable, Optional

import numpy as np
import pandas as pd
from PySide6.QtCore import QAbstractTableModel, Qt, QModelIndex


class PandasTableModel(QAbstractTableModel):
    def __init__(
        self,
        df: pd.DataFrame,
        parent=None,
        edit_callback: Optional[Callable[[Any, str, Any], None]] = None,
    ):
        """
        df            – DataFrame, который отображаем и редактируем
        edit_callback – функция (orig_index, column_name, new_value),
                        чтобы MainWindow мог обновить df_original
        """
        super().__init__(parent)
        self._df = df
        self._edit_callback = edit_callback

    # ---------- базовые методы ----------

    def rowCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self._df)

    def columnCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self._df.columns)

    def data(self, index: QModelIndex, role=Qt.DisplayRole) -> Any:
        if not index.isValid():
            return None

        if role == Qt.DisplayRole or role == Qt.EditRole:
            value = self._df.iat[index.row(), index.column()]

            if pd.isna(value):
                return ""

            # даты → дд.мм.рррр
            if isinstance(value, (pd.Timestamp, datetime.datetime, datetime.date)):
                return value.strftime("%d.%m.%Y")

            # bool → Так / Ні
            if isinstance(value, (bool, np.bool_)):
                return "Так" if value else "Ні"

            return str(value)

        return None

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role != Qt.DisplayRole:
            return None

        if orientation == Qt.Horizontal:
            return str(self._df.columns[section])
        else:
            return str(section + 1)

    # ---------- редактирование ----------

    def flags(self, index: QModelIndex):
        if not index.isValid():
            return Qt.ItemIsEnabled
        return (
            Qt.ItemIsEnabled
            | Qt.ItemIsSelectable
            | Qt.ItemIsEditable  # разрешаем редактирование
        )

    def setData(self, index: QModelIndex, value, role=Qt.EditRole) -> bool:
        if role != Qt.EditRole or not index.isValid():
            return False

        row = index.row()
        col = index.column()
        col_name = self._df.columns[col]
        series = self._df[col_name]

        text = str(value).strip()

        # Пытаемся привести к правильному типу
        new_val: Any = text

        try:
            if pd.api.types.is_datetime64_any_dtype(series):
                if text == "":
                    new_val = pd.NaT
                else:
                    # ожидаем формат дд.мм.рррр
                    new_val = pd.to_datetime(text, format="%d.%m.%Y")
            elif pd.api.types.is_bool_dtype(series):
                if text == "":
                    new_val = False
                else:
                    t = text.lower()
                    new_val = t in ("так", "true", "1", "yes", "y", "+")
            elif pd.api.types.is_numeric_dtype(series):
                if text == "":
                    new_val = np.nan
                else:
                    try:
                        new_val = int(text)
                    except ValueError:
                        new_val = float(text)
            else:
                # обычный текст
                new_val = text
        except Exception:
            # если что-то не так – оставляем как текст
            new_val = text

        # Обновляем текущий df (отфильтрованный)
        self._df.iat[row, col] = new_val

        # Коллбек в MainWindow, чтобы обновить df_original
        if self._edit_callback is not None:
            orig_index = self._df.index[row]  # индекс в оригинальном df
            self._edit_callback(orig_index, col_name, new_val)

        self.dataChanged.emit(index, index, [Qt.DisplayRole, Qt.EditRole])
        return True

    # ---------- обновление целого df ----------

    def update_df(self, df: pd.DataFrame):
        """Обновить данные в таблице (после фильтрации/редактирования)."""
        self.beginResetModel()
        self._df = df
        self.endResetModel()

    @property
    def df(self) -> pd.DataFrame:
        return self._df