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
        expiring_by5_indices: Set[Any] | None = None,   # жёлтая клетка в 5-й колонке
        expired_indices: Set[Any] | None = None,        # красный рядок (5 кол.)
        duplicate_indices: Set[Any] | None = None,      # синий рядок (дубликаты ПІБ)
        ors_warning_indices: Set[Any] | None = None,    # жёлтые 7–8 (до 20 діб)
        ors_overdue_indices: Set[Any] | None = None,    # красные 7–8 (>20 діб)
        col5_name: str | None = None,
        col7_name: str | None = None,
        col8_name: str | None = None,
    ):
        super().__init__()
        self.df = df
        self.edit_callback = edit_callback

        # Наборы индексов для подсветки
        self.expiring_by5_indices = expiring_by5_indices or set()
        self.expired_indices = expired_indices or set()
        self.duplicate_indices = duplicate_indices or set()
        self.ors_warning_indices = ors_warning_indices or set()
        self.ors_overdue_indices = ors_overdue_indices or set()

        # Семантические имена колонок
        self.col5_name = col5_name
        self.col7_name = col7_name
        self.col8_name = col8_name

    # ----- service -----

    def update_df(
        self,
        df: pd.DataFrame,
        expiring_by5_indices: Set[Any] | None = None,
        expired_indices: Set[Any] | None = None,
        duplicate_indices: Set[Any] | None = None,
        ors_warning_indices: Set[Any] | None = None,
        ors_overdue_indices: Set[Any] | None = None,
        col5_name: str | None = None,
        col7_name: str | None = None,
        col8_name: str | None = None,
    ):
        """
        Полное обновление DataFrame и наборов индексов подсветки.
        """
        self.beginResetModel()
        self.df = df

        if expiring_by5_indices is not None:
            self.expiring_by5_indices = expiring_by5_indices
        if expired_indices is not None:
            self.expired_indices = expired_indices
        if duplicate_indices is not None:
            self.duplicate_indices = duplicate_indices
        if ors_warning_indices is not None:
            self.ors_warning_indices = ors_warning_indices
        if ors_overdue_indices is not None:
            self.ors_overdue_indices = ors_overdue_indices

        if col5_name is not None:
            self.col5_name = col5_name
        if col7_name is not None:
            self.col7_name = col7_name
        if col8_name is not None:
            self.col8_name = col8_name

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

        # ----- Текст -----
        if role == Qt.DisplayRole:
            if pd.isna(value):
                return ""
            return str(value)

        if role == Qt.EditRole:
            return "" if pd.isna(value) else str(value)

        # ----- Фон -----
        if role == Qt.BackgroundRole:
            # 0) Видалені – весь рядок сірий
            if "is_deleted" in self.df.columns and bool(self.df.at[orig_index, "is_deleted"]):
                return QColor(230, 230, 230)

            # 1) Прострочено строк запобіжного заходу (5 кол.) – ВЕСЬ рядок червоний
            if orig_index in self.expired_indices:
                return QColor(255, 190, 190)

            # 2) Не заведено ОРС (понад 20 діб) – червоні КЛІТИНКИ 7–8 колонок
            if (
                self.col7_name is not None
                and self.col8_name is not None
                and orig_index in self.ors_overdue_indices
                and col_name in (self.col7_name, self.col8_name)
            ):
                return QColor(255, 200, 200)

            # 3) Строк запобіжного заходу спливає (10 діб до 6 міс.) – жовта КЛІТИНКА 5-ї колонки
            if (
                self.col5_name is not None
                and orig_index in self.expiring_by5_indices
                and col_name == self.col5_name
            ):
                return QColor(255, 245, 200)

            # 4) Не заведено ОРС (до 20 діб) – жовті КЛІТИНКИ 7–8 колонок
            if (
                self.col7_name is not None
                and self.col8_name is not None
                and orig_index in self.ors_warning_indices
                and col_name in (self.col7_name, self.col8_name)
            ):
                return QColor(255, 245, 200)

            # 5) Дублікати ПІБ – ВЕСЬ рядок синій (мягкий)
            if orig_index in self.duplicate_indices:
                return QColor(200, 220, 255)

            # 6) Архів – ВЕСЬ рядок зелений
            if "is_archived" in self.df.columns and bool(self.df.at[orig_index, "is_archived"]):
                return QColor(210, 240, 210)

        # ----- Выравнивание -----
        if role == Qt.TextAlignmentRole:
            if pd.api.types.is_numeric_dtype(self.df[col_name]):
                return int(Qt.AlignRight | Qt.AlignVCenter)
            return int(Qt.AlignLeft | Qt.AlignVCenter)

        return None

    def headerData(self, section: int, orientation, role: int = Qt.DisplayRole):
        if role != Qt.DisplayRole:
            return None
        if orientation == Qt.Horizontal:
            return str(self.df.columns[section])
        return str(self.df.index[section])

    def flags(self, index: QModelIndex):
        if not index.isValid():
            return Qt.ItemIsEnabled
        flags = Qt.ItemIsSelectable | Qt.ItemIsEnabled

        col_name = self.df.columns[index.column()]
        # запрещаем редактировать служебные флаги напрямую
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

        # мягкая конвертация типов
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

        self.dataChanged.emit(index, index, [Qt.DisplayRole, Qt.EditRole, Qt.BackgroundRole])
        return True