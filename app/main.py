import sys
import os
import json
from pathlib import Path
from typing import Any, Set

import pandas as pd
from docx import Document
from docx.enum.section import WD_ORIENT

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QFileDialog,
    QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QListWidget, QTableView,
    QMessageBox, QComboBox, QLineEdit, QDialog,
    QFormLayout, QDialogButtonBox, QTabWidget,
    QAbstractItemView,
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QPixmap

from app.model import PandasTableModel
from app.load_test_data import load_test_df
from app.filters_core import FilterCondition, Operator, apply_filters

CONFIG_PATH = Path.home() / ".table_filter_engine.json"
STATE_PATH = Path.home() / ".table_filter_engine_state.pkl"
SERVICE_COLS = {"is_archived", "is_deleted"}


def resource_path(rel_path: str) -> Path:
    """Корректный путь к ресурсам и в dev, и в exe (PyInstaller)."""
    if hasattr(sys, "_MEIPASS"):
        base = Path(sys._MEIPASS)
    else:
        base = Path(__file__).resolve().parent
    return base / rel_path


# ============================================================
#                 ДИАЛОГ ДОБАВЛЕНИЯ СТРОКИ
# ============================================================

class AddRowDialog(QDialog):
    """
    Диалог добавления нового записи.
    Поля максимально приближены к реальным колонкам фінальної таблиці.
    """

    def __init__(self, prosecutors: list[str] | None = None, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Додати новий запис")
        self.setModal(True)

        prosecutors = prosecutors or []
        layout = QFormLayout(self)

        # 1. Прокуратура
        self.prosecutor_cb = QComboBox(self)
        self.prosecutor_cb.addItem("")
        for p in sorted(prosecutors):
            self.prosecutor_cb.addItem(str(p))
        layout.addRow("Прокуратура:", self.prosecutor_cb)

        # 2. № провадження / кваліфікація
        self.case_edit = QLineEdit(self)
        self.case_edit.setPlaceholderText("№ провадження, дата, кваліфікація, орган…")
        layout.addRow("№ провадження / кваліфікація:", self.case_edit)

        # 3. Фабула
        self.fabula_edit = QLineEdit(self)
        self.fabula_edit.setPlaceholderText("Коротка фабула…")
        layout.addRow("Фабула:", self.fabula_edit)

        # 4. ПІБ + дати
        self.pib_edit = QLineEdit(self)
        self.pib_edit.setPlaceholderText("Прізвище Ім'я По батькові")
        layout.addRow("ПІБ підозрюваного:", self.pib_edit)

        self.dob_edit = QLineEdit(self)
        self.dob_edit.setPlaceholderText("дд.мм.рррр")
        layout.addRow("Дата народження:", self.dob_edit)

        self.notice_date_edit = QLineEdit(self)
        self.notice_date_edit.setPlaceholderText("дд.мм.рррр")
        layout.addRow("Дата повідомлення підозри:", self.notice_date_edit)

        # 5. Запобіжний захід
        self.measure_edit = QLineEdit(self)
        self.measure_edit.setPlaceholderText("Тримання під вартою / застава / ухвала …")
        layout.addRow("Запобіжний захід:", self.measure_edit)

        # 6. Підстава, дата зупинення
        self.stop_edit = QLineEdit(self)
        self.stop_edit.setPlaceholderText("Підстава, дата зупинення…")
        layout.addRow("Зупинення розслідування:", self.stop_edit)

        # 7. Доручення / клопотання
        self.order_edit = QLineEdit(self)
        self.order_edit.setPlaceholderText("Дата, вих. №, слідчий, адресат…")
        layout.addRow("Доручення / клопотання:", self.order_edit)

        # 8. № ОРС
        self.ors_edit = QLineEdit(self)
        self.ors_edit.setPlaceholderText("№ ОРС, дата заведення, категорія, орган…")
        layout.addRow("№ ОРС:", self.ors_edit)

        # 9. Перетин кордону
        self.border_edit = QLineEdit(self)
        self.border_edit.setPlaceholderText("Так/Ні, дата отримання інформації…")
        layout.addRow("Перетин кордону:", self.border_edit)

        # 10. Адмін. відповідальність
        self.admin_edit = QLineEdit(self)
        self.admin_edit.setPlaceholderText("Так/Ні, стаття, дата…")
        layout.addRow("Адмін. відповідальність:", self.admin_edit)

        # 11. Міжнародний розшук / Інтерпол
        self.interpol_edit = QLineEdit(self)
        self.interpol_edit.setPlaceholderText("Дата оголошення, № картки Інтерполу…")
        layout.addRow("Міжнародний розшук:", self.interpol_edit)

        # Кнопки
        btn_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, self)
        btn_box.accepted.connect(self.accept)
        btn_box.rejected.connect(self.reject)
        layout.addRow(btn_box)

    def get_data(self) -> dict[str, str]:
        return {
            "prosecutor": self.prosecutor_cb.currentText().strip(),
            "case_info": self.case_edit.text().strip(),
            "fabula": self.fabula_edit.text().strip(),
            "pib": self.pib_edit.text().strip(),
            "dob": self.dob_edit.text().strip(),
            "notice_date": self.notice_date_edit.text().strip(),
            "measure": self.measure_edit.text().strip(),
            "stop_info": self.stop_edit.text().strip(),
            "order_info": self.order_edit.text().strip(),
            "ors_info": self.ors_edit.text().strip(),
            "border_info": self.border_edit.text().strip(),
            "admin_info": self.admin_edit.text().strip(),
            "interpol_info": self.interpol_edit.text().strip(),
        }


# ============================================================
#                      ГЛАВНОЕ ОКНО
# ============================================================

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Table Filter Engine")
        self.resize(1500, 900)

        self.df_original: pd.DataFrame | None = None
        self.df_current: pd.DataFrame | None = None

        self.conditions: list[FilterCondition] = []
        self.global_search_text: str = ""
        self.expiring_indices: Set[Any] = set()
        self.duplicate_indices: Set[Any] = set()
        self.show_only_expiring: bool = False
        self.view_mode: str = "main"  # main / archive / deleted

        self.current_file_path: str | None = None

        self._init_ui()
        self._load_last_state_or_file()

    # --------------------------------------------------------
    #                    ИНИЦИАЛИЗАЦИЯ UI
    # --------------------------------------------------------

    def _init_ui(self):
        central = QWidget()
        root = QVBoxLayout(central)
        root.setContentsMargins(5, 5, 5, 5)

        # Верхняя панель
        top = QHBoxLayout()

        self.btn_load = QPushButton("📂 Відкрити")
        self.btn_load.clicked.connect(self.open_file)
        top.addWidget(self.btn_load)

        self.btn_add = QPushButton("➕ Додати рядок")
        self.btn_add.clicked.connect(self.add_row)
        self.btn_add.setEnabled(False)
        top.addWidget(self.btn_add)

        self.btn_export = QPushButton("💾 Експорт")
        self.btn_export.clicked.connect(self.export_file)
        self.btn_export.setEnabled(False)
        top.addWidget(self.btn_export)

        top.addStretch()

        top.addWidget(QLabel("Глобальний пошук:"))
        self.ed_search = QLineEdit()
        self.ed_search.setPlaceholderText("Пошук по всіх стовпцях…")
        self.ed_search.textChanged.connect(self.on_global_search)
        self.ed_search.setEnabled(False)
        top.addWidget(self.ed_search, stretch=2)

        # Вкладки режимов
        self.tab_mode = QTabWidget()
        self.tab_mode.addTab(QWidget(), "Основні")
        self.tab_mode.addTab(QWidget(), "Архів")
        self.tab_mode.addTab(QWidget(), "Видалені")
        self.tab_mode.currentChanged.connect(self.on_tab_changed)
        self.tab_mode.setTabPosition(QTabWidget.North)
        top.addWidget(self.tab_mode)

        root.addLayout(top)

        # Центральная часть
        main = QHBoxLayout()

        # Левая панель
        left = QVBoxLayout()
        left.setAlignment(Qt.AlignTop)

        # Фильтр по прокуратуре
        lbl_p = QLabel("Фільтр по прокуратурі:")
        lbl_p.setStyleSheet("font-weight: bold;")
        left.addWidget(lbl_p)

        self.cb_prosecutor = QComboBox()
        self.cb_prosecutor.addItem("Усі прокуратури")
        self.cb_prosecutor.currentIndexChanged.connect(self.apply_all_filters)
        self.cb_prosecutor.setEnabled(False)
        left.addWidget(self.cb_prosecutor)

        left.addSpacing(15)

        # Фильтр по столбцу
        lbl_c = QLabel("Фільтр по стовпцю:")
        lbl_c.setStyleSheet("font-weight: bold;")
        left.addWidget(lbl_c)

        self.cb_column = QComboBox()
        self.cb_column.setEnabled(False)
        self.cb_column.currentIndexChanged.connect(self.on_column_changed)
        left.addWidget(self.cb_column)

        self.cb_operator = QComboBox()
        self.cb_operator.addItems(["містить", "дорівнює", "не дорівнює"])
        self.cb_operator.setEnabled(False)
        left.addWidget(self.cb_operator)

        self.ed_value = QLineEdit()
        self.ed_value.setPlaceholderText("Значення для фільтра…")
        self.ed_value.setEnabled(False)
        left.addWidget(self.ed_value)

        # Список возможных значений
        self.cb_value_choices = QComboBox()
        self.cb_value_choices.setVisible(False)
        self.cb_value_choices.currentIndexChanged.connect(self.on_value_choice_selected)
        left.addWidget(self.cb_value_choices)

        # Поля дат для диапазона
        self.ed_date_from = QLineEdit()
        self.ed_date_from.setVisible(False)
        left.addWidget(self.ed_date_from)

        self.ed_date_to = QLineEdit()
        self.ed_date_to.setVisible(False)
        left.addWidget(self.ed_date_to)

        self.btn_add_condition = QPushButton("Додати умову")
        self.btn_add_condition.clicked.connect(self.add_condition_from_ui)
        self.btn_add_condition.setEnabled(False)
        left.addWidget(self.btn_add_condition)

        left.addSpacing(10)

        lbl_curr = QLabel("Поточні умови:")
        left.addWidget(lbl_curr)

        self.list_conditions = QListWidget()
        left.addWidget(self.list_conditions)

        self.btn_remove_condition = QPushButton("🗑 Видалити обрану умову")
        self.btn_remove_condition.clicked.connect(self.remove_selected_condition)
        self.btn_remove_condition.setEnabled(False)
        left.addWidget(self.btn_remove_condition)

        self.btn_clear_conditions = QPushButton("❌ Очистити всі умови")
        self.btn_clear_conditions.clicked.connect(self.clear_conditions)
        self.btn_clear_conditions.setEnabled(False)
        left.addWidget(self.btn_clear_conditions)

        # Кнопка "Показати строки зі строком, що спливає"
        self.btn_show_expiring = QPushButton("Показати строки зі строком, що спливає")
        self.btn_show_expiring.setEnabled(False)
        self.btn_show_expiring.setCheckable(True)
        self.btn_show_expiring.toggled.connect(self.on_toggle_show_expiring)
        left.addWidget(self.btn_show_expiring)

        # Кнопка: проверка дублей
        self.btn_check_duplicates = QPushButton("Перевірити дублікати")
        self.btn_check_duplicates.setEnabled(False)
        self.btn_check_duplicates.clicked.connect(self.on_check_duplicates)
        left.addWidget(self.btn_check_duplicates)

        left.addSpacing(10)

        # Операции с строками
        lbl_ops = QLabel("Операції з рядками (за виділенням):")
        lbl_ops.setStyleSheet("font-weight: bold;")
        left.addWidget(lbl_ops)

        self.btn_to_archive = QPushButton("В архів")
        self.btn_to_archive.clicked.connect(self.move_selected_to_archive)
        self.btn_to_archive.setEnabled(False)
        left.addWidget(self.btn_to_archive)

        self.btn_from_archive = QPushButton("З архіву")
        self.btn_from_archive.clicked.connect(self.move_selected_from_archive)
        self.btn_from_archive.setEnabled(False)
        left.addWidget(self.btn_from_archive)

        self.btn_delete_rows = QPushButton("Видалити")
        self.btn_delete_rows.clicked.connect(self.delete_selected_rows)
        self.btn_delete_rows.setEnabled(False)
        left.addWidget(self.btn_delete_rows)

        self.btn_restore_rows = QPushButton("Відновити")
        self.btn_restore_rows.clicked.connect(self.restore_selected_rows)
        self.btn_restore_rows.setEnabled(False)
        left.addWidget(self.btn_restore_rows)

        self.list_conditions.itemDoubleClicked.connect(
            lambda _: self.remove_selected_condition()
        )

        # Таблица
        self.table_view = QTableView()
        self.table_view.setAlternatingRowColors(True)
        self.table_view.horizontalHeader().setStretchLastSection(True)
        self.table_view.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table_view.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.table_view.setEditTriggers(
            QAbstractItemView.SelectedClicked
            | QAbstractItemView.DoubleClicked
            | QAbstractItemView.EditKeyPressed
        )

        # Левая панель уже: 1 к 6
        main.addLayout(left, 1)
        main.addWidget(self.table_view, 6)

        root.addLayout(main)

        # Нижний правый угол — логотип + копирайт
        footer = QHBoxLayout()
        footer.setSpacing(8)
        footer.addStretch()

        logo_label = QLabel()
        logo_path = resource_path("assets/national.png")
        if logo_path.exists():
            pm = QPixmap(str(logo_path))
            if not pm.isNull():
                pm = pm.scaledToHeight(69, Qt.SmoothTransformation)
                logo_label.setPixmap(pm)
        footer.addWidget(logo_label)

        copyright_label = QLabel("© Cyberpolice")
        copyright_label.setStyleSheet("color: rgba(255,255,255,150); font-size: 11px;")
        footer.addWidget(copyright_label)

        root.addLayout(footer)
        self.setCentralWidget(central)

    # --------------------------------------------------------
    #                    СЛУЖЕБНЫЕ МЕТОДЫ
    # --------------------------------------------------------

    def _is_date_like_column(self, series: pd.Series) -> bool:
        """Определяем, похож ли столбец на дату."""
        if pd.api.types.is_datetime64_any_dtype(series):
            return True
        try:
            return series.astype(str).str.contains(r"\d{2}\.\d{2}\.\d{4}").any()
        except Exception:
            return False

    def _save_last_file(self, path: str):
        try:
            CONFIG_PATH.write_text(
                json.dumps({"last_file": path}, ensure_ascii=False),
                encoding="utf-8",
            )
        except Exception:
            pass

    def _save_state(self):
        if self.df_original is None:
            return
        try:
            self.df_original.to_pickle(STATE_PATH)
        except Exception:
            pass

    def _load_last_state_or_file(self):
        # Сначала пробуем поднять состояние из pickle
        if STATE_PATH.exists():
            try:
                df = pd.read_pickle(STATE_PATH)
                self.current_file_path = None
                self._setup_dataframe(df, show_message=False)
                return
            except Exception:
                pass

        # Если состояния нет — пробуем последний файл
        if not CONFIG_PATH.exists():
            return
        try:
            data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
            last_path = data.get("last_file")
            if last_path and os.path.exists(last_path):
                self.load_table_from_path(last_path, show_message=False)
        except Exception:
            pass

    # --------------------------------------------------------
    #      ЕДИНОЕ место, где мы привязываем df к UI
    # --------------------------------------------------------

    def _setup_dataframe(self, df: pd.DataFrame, show_message: bool):
        self.df_original = df
        self.df_current = df.copy()

        self.recalc_expiring_marks(show_popup=show_message)
        self.recalc_duplicate_marks(show_popup=show_message)

        model = PandasTableModel(
            self.df_current,
            edit_callback=self.on_cell_edited,
            expiring_indices=self.expiring_indices,
            duplicate_indices=self.duplicate_indices,
        )
        self.table_view.setModel(model)
        self.hide_service_columns()

        # Прокуратуры
        self.cb_prosecutor.setEnabled(True)
        self.cb_prosecutor.clear()
        self.cb_prosecutor.addItem("Усі прокуратури")
        if "Прокуратура" in df.columns:
            for p in sorted(df["Прокуратура"].dropna().unique()):
                self.cb_prosecutor.addItem(str(p))

        # Столбцы без служебных
        self.cb_column.setEnabled(True)
        self.cb_column.clear()
        for col in df.columns:
            if col not in SERVICE_COLS:
                self.cb_column.addItem(col)

        self.cb_operator.setEnabled(True)
        self.ed_value.setEnabled(True)
        self.btn_add_condition.setEnabled(True)
        self.btn_clear_conditions.setEnabled(True)
        self.btn_remove_condition.setEnabled(True)

        self.btn_add.setEnabled(True)
        self.btn_export.setEnabled(True)
        self.ed_search.setEnabled(True)
        self.btn_show_expiring.setEnabled(bool(self.expiring_indices))
        self.btn_check_duplicates.setEnabled(True)

        self.conditions.clear()
        self.list_conditions.clear()
        self.global_search_text = ""
        self.ed_search.clear()
        self.show_only_expiring = False
        self.btn_show_expiring.setChecked(False)
        self.view_mode = "main"
        self.tab_mode.setCurrentIndex(0)
        self.update_action_buttons_state()

        self.on_column_changed(self.cb_column.currentIndex())
        self._save_state()

    # --------------------------------------------------------
    #                    ЗАГРУЗКА ТАБЛИЦЫ
    # --------------------------------------------------------

    def open_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Вибрати файл реєстру",
            "",
            "Таблиці (*.csv *.xlsx *.xls *.docx);;Усі файли (*)"
        )
        if not path:
            return
        self.load_table_from_path(path, show_message=True)

    def load_table_from_path(self, path: str, show_message: bool = True):
        try:
            df = load_test_df(path)

            if "is_archived" not in df.columns:
                df["is_archived"] = False
            if "is_deleted" not in df.columns:
                df["is_deleted"] = False

            self.current_file_path = path
            self._save_last_file(path)

            self._setup_dataframe(df, show_message=show_message)

            if show_message:
                QMessageBox.information(self, "OK", f"Файл завантажено:\n{path}")
        except Exception as e:
            QMessageBox.critical(self, "Помилка завантаження", str(e))

    def hide_service_columns(self):
        model = self.table_view.model()
        if not isinstance(model, PandasTableModel):
            return
        df = model.df
        for name in ("is_archived", "is_deleted"):
            if name in df.columns:
                idx = df.columns.get_loc(name)
                self.table_view.setColumnHidden(idx, True)

    # --------------------------------------------------------
    #              ПЕРЕСЧЁТ "ИСТЕКАЮЩИХ" СРОКОВ
    # --------------------------------------------------------

    def recalc_expiring_marks(self, show_popup: bool = True):
        """
        Отмечаем строки, где:
        - в колонке с запобіжним заходом дата "до" истекает ≤10 дней;
        - в колонке с ОРС от даты заведення прошло 0–20 дней.
        """
        self.expiring_indices = set()
        if self.df_original is None:
            return

        df = self.df_original
        today = pd.Timestamp.today().normalize()

        # Колонка 5
        col5 = next(
            (c for c in df.columns if "Запобіжн" in str(c) or "ухвала про дозвіл" in str(c)),
            None,
        )
        if col5:
            ser5 = df[col5].astype(str)
            matches5 = ser5.str.extractall(r"(\d{2}\.\d{2}\.\d{4})")
            if not matches5.empty:
                last_dates_str = matches5.groupby(level=0)[0].last()
                dates5 = pd.to_datetime(last_dates_str, format="%d.%m.%Y", errors="coerce")
                delta5 = (dates5 - today).dt.days
                idxs5 = dates5.index[(delta5 >= 0) & (delta5 <= 10)]
                self.expiring_indices.update(idxs5.tolist())

        # Колонка 8: № ОРС
        col8 = next(
            (c for c in df.columns if "№ОРС" in str(c) or "№ ОРС" in str(c)),
            None,
        )
        if col8:
            ser8 = df[col8].astype(str)
            first_dates_str = ser8.str.extract(r"(\d{2}\.\d{2}\.\d{4})")[0]
            dates8 = pd.to_datetime(first_dates_str, format="%d.%m.%Y", errors="coerce")
            delta8 = (today - dates8).dt.days
            idxs8 = dates8.index[(delta8 >= 0) & (delta8 <= 20)]
            self.expiring_indices.update(idxs8.tolist())

        if show_popup and self.expiring_indices:
            QMessageBox.warning(
                self,
                "Увага",
                f"Є {len(self.expiring_indices)} запис(ів) зі строком, що спливає.",
            )

    # --------------------------------------------------------
    #                ПОИСК ДУБЛИКАТОВ
    # --------------------------------------------------------

    def recalc_duplicate_marks(self, show_popup: bool = True):
        """
        Дубликаты ищем по полному совпадению ПІБ (часть до первой запятой)
        в колонке, где в названии есть 'ПІБ'.
        Учитываем только записи, у которых is_deleted == False.
        """
        old_count = len(self.duplicate_indices)
        self.duplicate_indices = set()

        if self.df_original is None:
            return

        df = self.df_original

        pib_col = next((c for c in df.columns if "ПІБ" in str(c)), None)
        if pib_col is None:
            return

        # Учитываем только не удалённые строки
        if "is_deleted" in df.columns:
            active_mask = df["is_deleted"] == False
        else:
            active_mask = pd.Series(True, index=df.index)

        if not active_mask.any():
            return

        full_series = df.loc[active_mask, pib_col].astype(str)
        name_series = full_series.str.split(",", n=1).str[0].str.strip()

        valid = name_series != ""
        name_valid = name_series[valid]
        if name_valid.empty:
            return

        counts = name_valid.value_counts()
        dup_names = set(counts[counts > 1].index)
        if not dup_names:
            return

        mask_dups = name_series.isin(dup_names)
        idxs = name_series.index[mask_dups].tolist()
        self.duplicate_indices.update(idxs)

        if show_popup and len(self.duplicate_indices) > old_count:
            QMessageBox.warning(
                self,
                "Дублікати",
                f"Виявлено {len(self.duplicate_indices)} запис(ів)-дублікат(ів) "
                f"(за повним збігом ПІБ).",
            )

    def on_check_duplicates(self):
        """Обработчик кнопки 'Перевірити дублікати'."""
        if self.df_original is None:
            QMessageBox.information(self, "Дублікати", "Немає завантаженої таблиці.")
            return
        self.recalc_duplicate_marks(show_popup=True)
        self.apply_all_filters()

    # --------------------------------------------------------
    #                   ГЛОБАЛЬНЫЙ ПОИСК
    # --------------------------------------------------------

    def on_global_search(self, text: str):
        self.global_search_text = text.strip()
        self.apply_all_filters()

    # --------------------------------------------------------
    #           ПЕРЕКЛЮЧЕНИЕ РЕЖИМА ВВОДА ДЛЯ СТОЛБЦОВ
    # --------------------------------------------------------

    def on_column_changed(self, index: int):
        if self.df_original is None or index < 0:
            return

        column = self.cb_column.itemText(index)
        if not column:
            return

        series = self.df_original[column]
        is_date_like = self._is_date_like_column(series)

        self.cb_operator.setVisible(True)
        self.ed_value.setVisible(True)

        if is_date_like:
            self.ed_date_from.setVisible(True)
            self.ed_date_to.setVisible(True)
            self.ed_date_from.setPlaceholderText("з дд.мм.рррр (можна не заповнювати)")
            self.ed_date_to.setPlaceholderText("по дд.мм.рррр (можна не заповнювати)")
        else:
            self.ed_date_from.setVisible(False)
            self.ed_date_to.setVisible(False)

        self.ed_date_from.clear()
        self.ed_date_to.clear()

        uniques = series.dropna().unique()
        if len(uniques) <= 50 or column in ("Стаття_ККУ", "Категорія_розшуку"):
            self.cb_value_choices.setVisible(True)
            self.cb_value_choices.clear()
            self.cb_value_choices.addItem("— оберіть значення —")
            for val in sorted(map(str, uniques)):
                self.cb_value_choices.addItem(val)
        else:
            self.cb_value_choices.setVisible(False)

    def on_value_choice_selected(self, index: int):
        if index <= 0:
            return
        text = self.cb_value_choices.currentText()
        self.ed_value.setText(text)

    # --------------------------------------------------------
    #                 ДОБАВЛЕНИЕ УСЛОВИЙ ФИЛЬТРА
    # --------------------------------------------------------

    def add_condition_from_ui(self):
        if self.df_original is None:
            return

        column = self.cb_column.currentText()
        if not column:
            return

        series = self.df_original[column]
        is_date_like = self._is_date_like_column(series)

        # Диапазон дат
        if is_date_like:
            from_text = self.ed_date_from.text().strip()
            to_text = self.ed_date_to.text().strip()

            if from_text or to_text:
                def parse_date(txt: str):
                    if not txt:
                        return None
                    try:
                        return pd.to_datetime(txt, format="%d.%m.%Y", dayfirst=True)
                    except Exception:
                        QMessageBox.warning(
                            self,
                            "Невірний формат дати",
                            "Використовуйте формат дд.мм.рррр (наприклад, 05.01.2025).",
                        )
                        raise

                try:
                    d_from = parse_date(from_text)
                    d_to = parse_date(to_text)
                except Exception:
                    return

                cond = FilterCondition(
                    column=column,
                    operator=Operator.RANGE,
                    value=(d_from, d_to),
                )
                self.conditions.append(cond)

                label_from = from_text or "…"
                label_to = to_text or "…"
                self.list_conditions.addItem(f"{column}: {label_from} — {label_to}")

                self.ed_date_from.clear()
                self.ed_date_to.clear()
                self.apply_all_filters()
                return

        # Текстовый фильтр
        op_text = self.cb_operator.currentText()
        raw_value = self.ed_value.text().strip()
        if not op_text or not raw_value:
            return

        if op_text == "містить":
            operator = Operator.CONTAINS
        elif op_text == "дорівнює":
            operator = Operator.EQUALS
        else:
            operator = Operator.NOT_EQUALS

        value: Any = raw_value

        try:
            if pd.api.types.is_bool_dtype(series):
                v = raw_value.lower()
                if v in ("так", "true", "1"):
                    value = True
                elif v in ("ні", "false", "0", "нет", "no"):
                    value = False
            elif pd.api.types.is_datetime64_any_dtype(series):
                value = pd.to_datetime(raw_value, format="%d.%m.%Y", dayfirst=True)
            elif pd.api.types.is_numeric_dtype(series):
                try:
                    value = int(raw_value)
                except ValueError:
                    value = float(raw_value)
        except Exception:
            value = raw_value

        cond = FilterCondition(column=column, operator=operator, value=value)
        self.conditions.append(cond)
        self.list_conditions.addItem(f"{column} {op_text} {raw_value}")

        self.ed_value.clear()
        self.apply_all_filters()

    def remove_selected_condition(self):
        idx = self.list_conditions.currentRow()
        if idx < 0 or idx >= len(self.conditions):
            return
        del self.conditions[idx]
        self.list_conditions.takeItem(idx)
        self.apply_all_filters()

    def clear_conditions(self):
        self.conditions.clear()
        self.list_conditions.clear()
        self.apply_all_filters()

    # --------------------------------------------------------
    #                   ВКЛАДКИ (РЕЖИМ ПРОСМОТРА)
    # --------------------------------------------------------

    def on_tab_changed(self, index: int):
        if index == 0:
            self.view_mode = "main"
        elif index == 1:
            self.view_mode = "archive"
        else:
            self.view_mode = "deleted"
        self.update_action_buttons_state()
        self.apply_all_filters()

    def update_action_buttons_state(self):
        if self.view_mode == "main":
            self.btn_to_archive.setEnabled(True)
            self.btn_from_archive.setEnabled(False)
            self.btn_delete_rows.setEnabled(True)
            self.btn_restore_rows.setEnabled(False)
        elif self.view_mode == "archive":
            self.btn_to_archive.setEnabled(False)
            self.btn_from_archive.setEnabled(True)
            self.btn_delete_rows.setEnabled(True)
            self.btn_restore_rows.setEnabled(False)
        else:  # deleted
            self.btn_to_archive.setEnabled(False)
            self.btn_from_archive.setEnabled(False)
            self.btn_delete_rows.setEnabled(False)
            self.btn_restore_rows.setEnabled(True)

    # --------------------------------------------------------
    #                  ПРИМЕНЕНИЕ ФИЛЬТРОВ
    # --------------------------------------------------------

    def apply_all_filters(self):
        if self.df_original is None:
            return

        df = self.df_original.copy()

        # 1) условия
        if self.conditions:
            df = apply_filters(df, self.conditions)

        # 2) прокуратура
        pros = self.cb_prosecutor.currentText()
        if pros and pros != "Усі прокуратури" and "Прокуратура" in df.columns:
            df = df[df["Прокуратура"] == pros]

        # 3) глобальный поиск
        if self.global_search_text:
            text = self.global_search_text
            mask = df.apply(
                lambda col: col.astype(str).str.contains(text, case=False, na=False),
                axis=0,
            ).any(axis=1)
            df = df[mask]

        # 4) режим просмотра
        if "is_deleted" in df.columns:
            if self.view_mode == "main":
                df = df[df["is_deleted"] == False]
            elif self.view_mode == "archive":
                df = df[(df["is_deleted"] == False) & (df["is_archived"] == True)]
            else:
                df = df[df["is_deleted"] == True]

        # 5) только строки со строком, що спливає
        if self.show_only_expiring and self.expiring_indices:
            df = df[df.index.isin(self.expiring_indices)]

        self.df_current = df

        model = self.table_view.model()
        if isinstance(model, PandasTableModel):
            model.update_df(
                self.df_current,
                expiring_indices=self.expiring_indices,
                duplicate_indices=self.duplicate_indices,
            )
        else:
            self.table_view.setModel(
                PandasTableModel(
                    self.df_current,
                    edit_callback=self.on_cell_edited,
                    expiring_indices=self.expiring_indices,
                    duplicate_indices=self.duplicate_indices,
                )
            )
        self.hide_service_columns()

    # --------------------------------------------------------
    #            СИНХРОНИЗАЦИЯ ПРАВОК В ТАБЛИЦЕ
    # --------------------------------------------------------

    def on_cell_edited(self, orig_index, column_name: str, new_value):
        if self.df_original is None:
            return
        if orig_index in self.df_original.index and column_name in self.df_original.columns:
            self.df_original.at[orig_index, column_name] = new_value

        # При любой осмысленной правке пересчитываем сроки;
        # дубликаты пользователь обновляет вручную кнопкой.
        if column_name not in ("is_archived", "is_deleted"):
            self.recalc_expiring_marks(show_popup=False)

        self._save_state()
        self.apply_all_filters()

    # --------------------------------------------------------
    #                  РАБОТА С ВЫДЕЛЕНИЕМ
    # --------------------------------------------------------

    def get_selected_indices(self) -> list[int]:
        if self.df_current is None:
            return []
        indices: set[int] = set()
        sel_model = self.table_view.selectionModel()
        if sel_model is not None:
            for idx in sel_model.selectedRows():
                try:
                    orig_index = self.df_current.index[idx.row()]
                    indices.add(orig_index)
                except Exception:
                    continue
        return list(indices)

    # --------------------------------------------------------
    #                     ДОБАВЛЕНИЕ СТРОКИ
    # --------------------------------------------------------

    def add_row(self):
        if self.df_original is None:
            return

        if "Прокуратура" in self.df_original.columns:
            prosecutors = sorted(self.df_original["Прокуратура"].dropna().unique())
        else:
            prosecutors = []

        dlg = AddRowDialog(prosecutors=prosecutors, parent=self)
        if dlg.exec() != QDialog.Accepted:
            return

        data = dlg.get_data()
        cols = list(self.df_original.columns)

        pib = data["pib"]
        dob = data["dob"]
        notice_date = data["notice_date"]
        pib_block = ", ".join([v for v in [pib, dob, notice_date] if v])

        # новый ID
        new_id = None
        if "ID" in cols:
            try:
                max_id = pd.to_numeric(self.df_original["ID"], errors="coerce").max()
                if pd.isna(max_id):
                    max_id = 0
                new_id = int(max_id) + 1
            except Exception:
                new_id = len(self.df_original) + 1

        row: dict[str, object] = {}

        for col in cols:
            text_col = str(col)

            if col == "ID" and new_id is not None:
                row[col] = new_id
            elif text_col == "Прокуратура":
                row[col] = data["prosecutor"]
            elif "№ кримінального провадження" in text_col:
                row[col] = data["case_info"]
            elif text_col.strip() == "Фабула":
                row[col] = data["fabula"]
            elif "ПІБ підозрюваного" in text_col:
                row[col] = pib_block
            elif "Запобіжний захід" in text_col:
                row[col] = data["measure"]
            elif "Підстава, дата зупинення" in text_col:
                row[col] = data["stop_info"]
            elif "Дата та вихідний № доручення" in text_col:
                row[col] = data["order_info"]
            elif "№ ОРС, дата заведення" in text_col:
                row[col] = data["ors_info"]
            elif "Наявність інформації про перетин кордону" in text_col:
                row[col] = data["border_info"]
            elif "Притягнення до адмін" in text_col:
                row[col] = data["admin_info"]
            elif "Дата оголошення у міжнародний розшук" in text_col:
                row[col] = data["interpol_info"]
            elif col == "is_archived":
                row[col] = False
            elif col == "is_deleted":
                row[col] = False
            else:
                row[col] = ""

        new_row_df = pd.DataFrame([row], columns=self.df_original.columns)
        self.df_original = pd.concat([self.df_original, new_row_df], ignore_index=True)

        self.recalc_expiring_marks(show_popup=False)
        self.recalc_duplicate_marks(show_popup=True)
        self._save_state()
        self.apply_all_filters()

    # --------------------------------------------------------
    #                     ОПЕРАЦИИ С СТРОКАМИ
    # --------------------------------------------------------

    def move_selected_to_archive(self):
        idxs = self.get_selected_indices()
        if not idxs:
            QMessageBox.information(self, "Архів", "Не вибрано жодного рядка.")
            return
        self.df_original.loc[idxs, "is_archived"] = True
        self._save_state()
        self.apply_all_filters()

    def move_selected_from_archive(self):
        idxs = self.get_selected_indices()
        if not idxs:
            QMessageBox.information(self, "Архів", "Не вибрано жодного рядка.")
            return
        self.df_original.loc[idxs, "is_archived"] = False
        self._save_state()
        self.apply_all_filters()

    def delete_selected_rows(self):
        idxs = self.get_selected_indices()
        if not idxs:
            QMessageBox.information(self, "Видалення", "Не вибрано жодного рядка.")
            return
        self.df_original.loc[idxs, "is_deleted"] = True

        # после удаления сразу пересчитываем дубликаты
        self.recalc_duplicate_marks(show_popup=False)

        self._save_state()
        self.apply_all_filters()

    def restore_selected_rows(self):
        idxs = self.get_selected_indices()
        if not idxs:
            QMessageBox.information(self, "Відновлення", "Не вибрано жодного рядка.")
            return
        self.df_original.loc[idxs, "is_deleted"] = False

        # после восстановления тоже пересчитываем дубликаты
        self.recalc_duplicate_marks(show_popup=False)

        self._save_state()
        self.apply_all_filters()

    # --------------------------------------------------------
    #            ПЕРЕКЛЮЧАТЕЛЬ "ПОКАЗАТИ СТРОКИ, ЩО СПЛИВАЮТЬ"
    # --------------------------------------------------------

    def on_toggle_show_expiring(self, checked: bool):
        self.show_only_expiring = checked
        self.apply_all_filters()

    # --------------------------------------------------------
    #                        ЭКСПОРТ
    # --------------------------------------------------------

    def _format_df_for_export(self, df: pd.DataFrame) -> pd.DataFrame:
        out = df.copy()
        for c in SERVICE_COLS:
            if c in out.columns:
                out = out.drop(columns=[c])
        for col in out.columns:
            if pd.api.types.is_datetime64_any_dtype(out[col]):
                out[col] = out[col].dt.strftime("%d.%m.%Y").fillna("")
            elif pd.api.types.is_bool_dtype(out[col]):
                out[col] = out[col].map({True: "Так", False: "Ні"})
        return out

    def export_file(self):
        if self.df_current is None or self.df_current.empty:
            QMessageBox.warning(self, "Експорт", "Немає даних для експорту.")
            return

        path, selected_filter = QFileDialog.getSaveFileName(
            self,
            "Зберегти результати фільтрації",
            "",
            "Word (*.docx);;Excel (*.xlsx);;CSV (*.csv)"
        )
        if not path:
            return

        try:
            df_out = self._format_df_for_export(self.df_current)

            if path.lower().endswith(".docx") or "Word" in selected_filter:
                doc = Document()
                section = doc.sections[0]
                section.orientation = WD_ORIENT.LANDSCAPE
                new_width, new_height = section.page_height, section.page_width
                section.page_width = new_width
                section.page_height = new_height

                table = doc.add_table(rows=1, cols=len(df_out.columns))
                table.style = "Table Grid"

                hdr_cells = table.rows[0].cells
                for j, col_name in enumerate(df_out.columns):
                    hdr_cells[j].text = str(col_name)

                for _, row in df_out.iterrows():
                    row_cells = table.add_row().cells
                    for j, col_name in enumerate(df_out.columns):
                        value = row[col_name]
                        row_cells[j].text = "" if pd.isna(value) else str(value)

                doc.save(path)

            elif path.lower().endswith(".xlsx") or "Excel" in selected_filter:
                df_out.to_excel(path, index=False)
            else:
                df_out.to_csv(path, index=False)

            QMessageBox.information(self, "Експорт", f"Файл збережено:\n{path}")
        except Exception as e:
            QMessageBox.critical(self, "Помилка експорту", str(e))


# ============================================================
#                  ТОЧКА ВХОДА
# ============================================================

def main():
    app = QApplication(sys.argv)

    app.setStyleSheet("""
        QWidget {
            background-color: #1e1e1e;
            color: #f0f0f0;
            font-size: 14px;
        }
        QPushButton {
            background-color: #333;
            border: 1px solid #555;
            padding: 6px 10px;
            border-radius: 4px;
        }
        QPushButton:hover {
            background-color: #444;
        }
        QTableView {
            gridline-color: #444;
            selection-background-color: #555;
        }
        QLineEdit {
            background-color: #2a2a2a;
            border: 1px solid #555;
            border-radius: 4px;
            padding: 4px;
        }
        QComboBox {
            background-color: #2a2a2a;
            border: 1px solid #555;
            border-radius: 4px;
            padding: 2px 4px;
        }
        QListWidget {
            background-color: #202020;
            border: 1px solid #444;
        }
    """)

    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()