from __future__ import annotations

from typing import Any, Set, Callable

import pandas as pd
from PySide6.QtCore import QAbstractTableModel, Qt, QModelIndex
from PySide6.QtGui import QColor


class PandasTableModel(QAbstractTableModel):
    def __init__(
        self,
        df: pd.DataFrame,
        edit_callback: Callable[[Any, str, Any], None] | None = None,
        expiring_indices: Set[Any] | None = None,
        expired_indices: Set[Any] | None = None,
        duplicate_indices: Set[Any] | None = None,
    ):
        super().__init__()
        self.df = df
        self.edit_callback = edit_callback
        self.expiring_indices = expiring_indices or set()
        self.expired_indices = expired_indices or set()
        self.duplicate_indices = duplicate_indices or set()

    # ----- service -----

    def update_df(
        self,
        df: pd.DataFrame,
        expiring_indices: Set[Any] | None = None,
        expired_indices: Set[Any] | None = None,
        duplicate_indices: Set[Any] | None = None,
    ):
        self.beginResetModel()
        self.df = df
        if expiring_indices is not None:
            self.expiring_indices = expiring_indices
        if expired_indices is not None:
            self.expired_indices = expired_indices
        if duplicate_indices is not None:
            self.duplicate_indices = duplicate_indices
        self.endResetModel()

    # ----- required overrides -----

    def rowCount(self, parent: QModelIndex = QModelIndex()) -> int:
        return 0 if parent.isValid() else len(self.df)

    def columnCount(self, parent: QModelIndex = QModelIndex()) -> int:
        return 0 if parent.isValid() else len(self.df.columns)

    def data(self, index: QModelIndex, role: int = Qt.DisplayRole):
        if not index.isValid():
            return None

        row = index.row()
        col = index.column()
        col_name = self.df.columns[col]
        orig_index = self.df.index[row]

        value = self.df.iat[row, col]

        if role == Qt.DisplayRole:
            if pd.isna(value):
                return ""
            return str(value)

        if role == Qt.EditRole:
            return "" if pd.isna(value) else str(value)

        if role == Qt.BackgroundRole:
            # базовый цвет (по умолчанию белый)
            color = None

            # deleted – слегка “затуманенный”
            if "is_deleted" in self.df.columns and bool(self.df.at[orig_index, "is_deleted"]):
                color = QColor(230, 230, 230)

            # archived – зелёный полупрозрачный
            if "is_archived" in self.df.columns and bool(self.df.at[orig_index, "is_archived"]):
                color = QColor(0, 140, 0, 40)

            # строки зі строком, що спливає (червоний)
            if orig_index in self.expiring_indices:
                color = QColor(180, 0, 0, 45)

            # прострочені (синій, ненавязчивый)
            if orig_index in self.expired_indices:
                color = QColor(40, 80, 200, 40)

            # дублікат ПІБ – жёлтый приоритетнее
            if orig_index in self.duplicate_indices:
                color = QColor(255, 215, 0, 80)

            if color is not None:
                return color

        if role == Qt.TextAlignmentRole:
            if pd.api.types.is_numeric_dtype(self.df[col_name]):
                return int(Qt.AlignRight | Qt.AlignVCenter)
            return int(Qt.AlignLeft | Qt.AlignVCenter)

        return None

    def headerData(self, section: int, orientation: Qt.Orientation, role: int = Qt.DisplayRole):
        if role != Qt.DisplayRole:
            return None
        if orientation == Qt.Horizontal:
            return str(self.df.columns[section])
        return str(self.df.index[section])

    def flags(self, index: QModelIndex):
        if not index.isValid():
            return Qt.ItemIsEnabled
        flags = Qt.ItemIsSelectable | Qt.ItemIsEnabled
        # запрещаем редактировать служебные флаги напрямую
        col_name = self.df.columns[index.column()]
        if col_name not in ("is_archived", "is_deleted"):
            flags |= Qt.ItemIsEditable
        return flags

    def setData(self, index: QModelIndex, value, role: int = Qt.EditRole):
        if role != Qt.EditRole or not index.isValid():
            return False

        row = index.row()
        col = index.column()
        col_name = self.df.columns[col]
        orig_index = self.df.index[row]

        # конвертация типов – максимально мягкая
        try:
            if pd.api.types.is_numeric_dtype(self.df[col_name]):
                if value == "":
                    new_val = pd.NA
                else:
                    try:
                        new_val = int(value)
                    except ValueError:
                        new_val = float(value)
            else:
                new_val = value
        except Exception:
            new_val = value

        self.df.iat[row, col] = new_val

        if self.edit_callback:
            self.edit_callback(orig_index, col_name, new_val)

        self.dataChanged.emit(index, index, [Qt.DisplayRole, Qt.EditRole])
        return True