from typing import Callable, Optional, Set, Any

import pandas as pd
from PySide6.QtCore import QAbstractTableModel, Qt, QModelIndex
from PySide6.QtGui import QColor, QBrush


class PandasTableModel(QAbstractTableModel):
    def __init__(
        self,
        df: pd.DataFrame,
        edit_callback: Optional[Callable[[Any, str, Any], None]] = None,
        expiring_indices: Optional[Set[Any]] = None,
        duplicate_indices: Optional[Set[Any]] = None,
    ):
        super().__init__()
        self.df = df
        self.edit_callback = edit_callback
        self.expiring_indices = expiring_indices or set()
        self.duplicate_indices = duplicate_indices or set()

    # --------------------------------------------------------
    #                 Обновление DataFrame
    # --------------------------------------------------------

    def update_df(
        self,
        df: pd.DataFrame,
        expiring_indices: Optional[Set[Any]] = None,
        duplicate_indices: Optional[Set[Any]] = None,
    ):
        self.beginResetModel()
        self.df = df
        if expiring_indices is not None:
            self.expiring_indices = expiring_indices
        if duplicate_indices is not None:
            self.duplicate_indices = duplicate_indices
        self.endResetModel()

    # --------------------------------------------------------
    #                  Базовые методы модели
    # --------------------------------------------------------

    def rowCount(self, parent: QModelIndex = QModelIndex()) -> int:
        if parent.isValid():
            return 0
        return len(self.df)

    def columnCount(self, parent: QModelIndex = QModelIndex()) -> int:
        if parent.isValid():
            return 0
        return len(self.df.columns)

    def headerData(self, section: int, orientation, role: int = Qt.DisplayRole):
        if role != Qt.DisplayRole:
            return None
        if orientation == Qt.Horizontal:
            try:
                return str(self.df.columns[section])
            except IndexError:
                return ""
        else:
            # можно вернуть индекс +1 или пусто
            return str(section + 1)

    # --------------------------------------------------------
    #                     Отображение данных
    # --------------------------------------------------------

    def data(self, index: QModelIndex, role: int = Qt.DisplayRole):
        if not index.isValid():
            return None

        row = index.row()
        col = index.column()

        if row < 0 or row >= len(self.df) or col < 0 or col >= len(self.df.columns):
            return None

        col_name = self.df.columns[col]
        orig_index = self.df.index[row]
        value = self.df.iloc[row, col]

        # Подсветка строк
        if role == Qt.BackgroundRole:
            # is_deleted / is_archived читаем из самой строки, если есть
            is_archived = bool(self.df.iloc[row][col_name] if col_name == "is_archived" else
                               (self.df.iloc[row].get("is_archived", False)
                                if "is_archived" in self.df.columns else False))
            is_deleted = bool(self.df.iloc[row][col_name] if col_name == "is_deleted" else
                              (self.df.iloc[row].get("is_deleted", False)
                               if "is_deleted" in self.df.columns else False))

            # Правила приоритетов:
            # 1) строки зі строком, що спливає – червоний
            # 2) видалені – сірі
            # 3) архівні – зелені
            # 4) дублікати – жовті
            if orig_index in self.expiring_indices:
                return QBrush(QColor(255, 0, 0, 60))
            if is_deleted:
                return QBrush(QColor(120, 120, 120, 80))
            if is_archived:
                return QBrush(QColor(0, 180, 0, 60))
            if orig_index in self.duplicate_indices:
                return QBrush(QColor(255, 255, 0, 80))

            return None

        # Цвет текста для видаленных (чуть бледнее)
        if role == Qt.ForegroundRole:
            if "is_deleted" in self.df.columns:
                if bool(self.df.iloc[row].get("is_deleted", False)):
                    return QBrush(QColor(200, 200, 200, 180))
            return None

        if role == Qt.DisplayRole or role == Qt.EditRole:
            # Красивый вывод bool: Так / Ні
            if isinstance(value, bool):
                if role == Qt.DisplayRole:
                    return "Так" if value else "Ні"
                else:
                    return value

            if pd.isna(value):
                return ""
            return str(value)

        return None

    # --------------------------------------------------------
    #                      Редактирование
    # --------------------------------------------------------

    def flags(self, index: QModelIndex):
        if not index.isValid():
            return Qt.NoItemFlags
        # всё редактируемое, кроме служебных колонок, если нужно — можно ограничить
        col_name = self.df.columns[index.column()]
        if col_name in ("is_archived", "is_deleted"):
            return Qt.ItemIsSelectable | Qt.ItemIsEnabled
        return Qt.ItemIsSelectable | Qt.ItemIsEnabled | Qt.ItemIsEditable

    def setData(self, index: QModelIndex, value, role: int = Qt.EditRole):
        if not index.isValid():
            return False

        row = index.row()
        col = index.column()
        col_name = self.df.columns[col]
        orig_index = self.df.index[row]

        if role == Qt.EditRole:
            # приведём boolы и числа к типу колонки, если получится
            series = self.df[col_name]
            new_val = value

            try:
                if pd.api.types.is_bool_dtype(series):
                    v = str(value).lower()
                    if v in ("так", "true", "1", "yes", "y"):
                        new_val = True
                    elif v in ("ні", "false", "0", "no", "n"):
                        new_val = False
                elif pd.api.types.is_numeric_dtype(series):
                    new_val = float(value)
                else:
                    new_val = value
            except Exception:
                new_val = value

            self.df.iat[row, col] = new_val
            self.dataChanged.emit(index, index, [Qt.DisplayRole, Qt.EditRole])

            if self.edit_callback:
                self.edit_callback(orig_index, col_name, new_val)

            return True

        return False