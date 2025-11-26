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
    QFormLayout, QDialogButtonBox
)
from PySide6.QtCore import Qt

from app.model import PandasTableModel
from app.load_test_data import load_test_df
from app.filters_core import FilterCondition, Operator, apply_filters


CONFIG_PATH = Path.home() / ".table_filter_engine.json"


# ============================================================
#                 ДІАЛОГ ДОДАВАННЯ РЯДКА
# ============================================================

class AddRowDialog(QDialog):
    """Діалог для додавання нового запису (базовий набір полів)."""

    def __init__(self, prosecutors: list[str], parent=None):
        super().__init__(parent)
        self.setWindowTitle("Додати новий запис")
        form = QFormLayout(self)

        # Прокуратура
        self.cb_prosecutor = QComboBox()
        self.cb_prosecutor.addItems(sorted(prosecutors) or ["—"])
        form.addRow("Прокуратура:", self.cb_prosecutor)

        # ПІБ
        self.ed_pib = QLineEdit()
        self.ed_pib.setPlaceholderText("Прізвище Ім'я По батькові")
        form.addRow("ПІБ:", self.ed_pib)

        # Дата народження
        self.ed_dob = QLineEdit()
        self.ed_dob.setPlaceholderText("дд.мм.рррр")
        form.addRow("Дата народження:", self.ed_dob)

        # Стаття
        self.ed_article = QLineEdit()
        form.addRow("Стаття ККУ / кваліфікація:", self.ed_article)

        # Запобіжний захід
        self.ed_measure = QLineEdit()
        form.addRow("Запобіжний захід:", self.ed_measure)

        # Виїзд за кордон
        self.cb_went = QComboBox()
        self.cb_went.addItems(["Ні", "Так"])
        form.addRow("Є виїзд за кордон:", self.cb_went)

        # Дата виїзду
        self.ed_depart = QLineEdit()
        self.ed_depart.setPlaceholderText("дд.мм.рррр або порожньо")
        form.addRow("Дата виїзду:", self.ed_depart)

        # Країна виїзду
        self.ed_country = QLineEdit()
        form.addRow("Країна виїзду:", self.ed_country)

        # Інтерпол
        self.cb_interpol = QComboBox()
        self.cb_interpol.addItems(["Ні", "Так"])
        form.addRow("Є Інтерпол:", self.cb_interpol)

        # Дата оголошення в розшук
        self.ed_interpol_date = QLineEdit()
        self.ed_interpol_date.setPlaceholderText("дд.мм.рррр або порожньо")
        form.addRow("Дата оголошення в розшук:", self.ed_interpol_date)

        # Примітка
        self.ed_note = QLineEdit()
        form.addRow("Примітка:", self.ed_note)

        buttons = QDialogButtonBox(
            QDialogButtonBox.Ok | QDialogButtonBox.Cancel,
            orientation=Qt.Horizontal
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        form.addRow(buttons)

    def get_data(self) -> dict:
        return {
            "Прокуратура": self.cb_prosecutor.currentText().strip(),
            "ПІБ": self.ed_pib.text().strip(),
            "Дата_нар": self.ed_dob.text().strip(),
            "Стаття_ККУ": self.ed_article.text().strip(),
            "Запобіжний_захід": self.ed_measure.text().strip(),
            "Є_виїзд_за_кордон": True if self.cb_went.currentText() == "Так" else False,
            "Дата_виїзду": self.ed_depart.text().strip(),
            "Країна_виїзду": self.ed_country.text().strip(),
            "Є_Інтерпол": True if self.cb_interpol.currentText() == "Так" else False,
            "Дата_оголошення_в_розшук": self.ed_interpol_date.text().strip(),
            "Примітка": self.ed_note.text().strip(),
        }


# ============================================================
#                      ГОЛОВНЕ ВІКНО
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
        self.highlight_indices: Set[Any] = set()
        self.show_only_expiring: bool = False

        self._init_ui()
        self._load_last_file()

    # --------------------------------------------------------
    #                    ІНІЦІАЛІЗАЦІЯ UI
    # --------------------------------------------------------

    def _init_ui(self):
        central = QWidget()
        root = QVBoxLayout(central)
        root.setContentsMargins(5, 5, 5, 5)

        # Верхня панель
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

        root.addLayout(top)

        # Центральна частина
        main = QHBoxLayout()

        # Ліва панель
        left = QVBoxLayout()
        left.setAlignment(Qt.AlignTop)

        # Фільтр по прокуратурі
        lbl_p = QLabel("Фільтр по прокуратурі:")
        lbl_p.setStyleSheet("font-weight: bold;")
        left.addWidget(lbl_p)

        self.cb_prosecutor = QComboBox()
        self.cb_prosecutor.addItem("Усі прокуратури")
        self.cb_prosecutor.currentIndexChanged.connect(self.apply_all_filters)
        self.cb_prosecutor.setEnabled(False)
        left.addWidget(self.cb_prosecutor)

        left.addSpacing(15)

        # Фільтр по стовпцю
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

        # Випадаючий список можливих значень
        self.cb_value_choices = QComboBox()
        self.cb_value_choices.setVisible(False)
        self.cb_value_choices.currentIndexChanged.connect(self.on_value_choice_selected)
        left.addWidget(self.cb_value_choices)

        # Поля дат для гнучкого діапазону
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

        self.list_conditions.itemDoubleClicked.connect(
            lambda _: self.remove_selected_condition()
        )

        # Таблиця справа
        self.table_view = QTableView()
        self.table_view.setAlternatingRowColors(True)
        self.table_view.horizontalHeader().setStretchLastSection(True)

        main.addLayout(left, 1)
        main.addWidget(self.table_view, 4)

        root.addLayout(main)
        self.setCentralWidget(central)

    # --------------------------------------------------------
    #                    ДОПОМІЖНІ МЕТОДИ
    # --------------------------------------------------------

    def _is_date_like_column(self, series: pd.Series) -> bool:
        """
        Визначає, чи можна вважати стовпець "датоподібним":
        - або це datetime64,
        - або в ньому є дата формату дд.мм.рррр (навіть у тексті).
        """
        if pd.api.types.is_datetime64_any_dtype(series):
            return True
        try:
            return series.astype(str).str.contains(r"\d{2}\.\d{2}\.\d{4}").any()
        except Exception:
            return False

    # --------------------------------------------------------
    #                   КОНФІГ (останній файл)
    # --------------------------------------------------------

    def _load_last_file(self):
        if not CONFIG_PATH.exists():
            return
        try:
            data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
            last_path = data.get("last_file")
            if last_path and os.path.exists(last_path):
                self.load_table_from_path(last_path, show_message=False)
        except Exception:
            pass

    def _save_last_file(self, path: str):
        try:
            CONFIG_PATH.write_text(
                json.dumps({"last_file": path}, ensure_ascii=False),
                encoding="utf-8",
            )
        except Exception:
            pass

    # --------------------------------------------------------
    #                    ЗАВАНТАЖЕННЯ ТАБЛИЦІ
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
            self.df_original = df
            self.df_current = df.copy()

            self.recalc_expiring_marks(show_popup=show_message)

            model = PandasTableModel(
                self.df_current,
                edit_callback=self.on_cell_edited,
                highlight_indices=self.highlight_indices,
            )
            self.table_view.setModel(model)

            # Прокуратури
            self.cb_prosecutor.setEnabled(True)
            self.cb_prosecutor.clear()
            self.cb_prosecutor.addItem("Усі прокуратури")
            if "Прокуратура" in df.columns:
                for p in sorted(df["Прокуратура"].dropna().unique()):
                    self.cb_prosecutor.addItem(str(p))

            # Стовпці
            self.cb_column.setEnabled(True)
            self.cb_column.clear()
            for col in df.columns:
                self.cb_column.addItem(col)

            self.cb_operator.setEnabled(True)
            self.ed_value.setEnabled(True)
            self.btn_add_condition.setEnabled(True)
            self.btn_clear_conditions.setEnabled(True)
            self.btn_remove_condition.setEnabled(True)

            self.btn_add.setEnabled(True)
            self.btn_export.setEnabled(True)
            self.ed_search.setEnabled(True)
            self.btn_show_expiring.setEnabled(bool(self.highlight_indices))

            self.conditions.clear()
            self.list_conditions.clear()
            self.global_search_text = ""
            self.ed_search.clear()
            self.show_only_expiring = False
            self.btn_show_expiring.setChecked(False)

            self.on_column_changed(self.cb_column.currentIndex())
            self._save_last_file(path)

            if show_message:
                QMessageBox.information(self, "OK", f"Файл завантажено:\n{path}")
        except Exception as e:
            QMessageBox.critical(self, "Помилка завантаження", str(e))

    # --------------------------------------------------------
    #              ПЕРЕРАХУНОК "СПЛИВАЮЧИХ" СТРОКІВ
    # --------------------------------------------------------

    def recalc_expiring_marks(self, show_popup: bool = True):
        """
        Позначає рядки, де:
        - у колонці із запобіжним заходом дата "до" спливає ≤10 днів;
        - у колонці з ОРС від дати заведення минуло 0–20 днів.
        """
        self.highlight_indices = set()
        if self.df_original is None:
            return

        df = self.df_original
        today = pd.Timestamp.today().normalize()

        # ----- Колонка 5: Запобіжний захід / ухвала про дозвіл ... -----
        col5 = next(
            (c for c in df.columns if "Запобіжн" in c or "ухвала про дозвіл" in c),
            None,
        )
        if col5:
            ser5 = df[col5].astype(str)
            # Вытаскиваем ВСЕ даты в строке, берём последнюю как "до"
            matches5 = ser5.str.extractall(r"(\d{2}\.\d{2}\.\d{4})")
            if not matches5.empty:
                last_dates_str = matches5.groupby(level=0)[0].last()
                dates5 = pd.to_datetime(
                    last_dates_str, format="%d.%m.%Y", errors="coerce"
                )
                delta5 = (dates5 - today).dt.days
                # до дати залишилося від 0 до 10 днів
                idxs5 = dates5.index[(delta5 >= 0) & (delta5 <= 10)]
                self.highlight_indices.update(idxs5.tolist())

        # ----- Колонка 8: № ОРС, дата заведення ... -----
        col8 = next(
            (c for c in df.columns if "№ОРС" in c or "№ ОРС" in c or "№ ОРС," in c),
            None,
        )
        if col8:
            ser8 = df[col8].astype(str)
            # Берём первую дату в строке как дату заведення
            first_dates_str = ser8.str.extract(r"(\d{2}\.\d{2}\.\d{4})")[0]
            dates8 = pd.to_datetime(
                first_dates_str, format="%d.%m.%Y", errors="coerce"
            )
            delta8 = (today - dates8).dt.days
            # від дати заведення минуло від 0 до 20 днів
            idxs8 = dates8.index[(delta8 >= 0) & (delta8 <= 20)]
            self.highlight_indices.update(idxs8.tolist())

        if show_popup and self.highlight_indices:
            QMessageBox.warning(
                self,
                "Увага",
                f"Є {len(self.highlight_indices)} запис(ів) зі строком, що спливає.",
            )

    # --------------------------------------------------------
    #                   ГЛОБАЛЬНИЙ ПОШУК
    # --------------------------------------------------------

    def on_global_search(self, text: str):
        self.global_search_text = text.strip()
        self.apply_all_filters()

    # --------------------------------------------------------
    #           ПЕРЕМИКАННЯ РЕЖИМУ ВВЕДЕННЯ ДЛЯ СТОВПЦІВ
    # --------------------------------------------------------

    def on_column_changed(self, index: int):
        if self.df_original is None or index < 0:
            return

        column = self.cb_column.itemText(index)
        if not column:
            return

        series = self.df_original[column]
        is_date_like = self._is_date_like_column(series)

        # В любом случае оставляем оператор + текст (чтобы можно было фильтровать по номеру)
        self.cb_operator.setVisible(True)
        self.ed_value.setVisible(True)

        if is_date_like:
            # Дополнительно показываем поля для диапазона дат
            self.ed_date_from.setVisible(True)
            self.ed_date_to.setVisible(True)
            self.ed_date_from.setPlaceholderText(
                "з дд.мм.рррр (можна не заповнювати)"
            )
            self.ed_date_to.setPlaceholderText(
                "по дд.мм.рррр (можна не заповнювати)"
            )
        else:
            self.ed_date_from.setVisible(False)
            self.ed_date_to.setVisible(False)

        self.ed_date_from.clear()
        self.ed_date_to.clear()

        # Випадаючий список можливих значень
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
    #                 ДОДАВАННЯ УМОВ ФІЛЬТРУ
    # --------------------------------------------------------

    def add_condition_from_ui(self):
        if self.df_original is None:
            return

        column = self.cb_column.currentText()
        if not column:
            return

        series = self.df_original[column]
        is_date_like = self._is_date_like_column(series)

        # --------- Сначала пробуем диапазон дат, если что-то введено ---------
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

        # --------- Если диапазон не задан — обычный текстовый фильтр ---------
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
    #           ПЕРЕМИКАЧ "ПОКАЗАТИ СТРОКИ ЗІ СТРОКОМ..."
    # --------------------------------------------------------

    def on_toggle_show_expiring(self, checked: bool):
        self.show_only_expiring = checked
        self.apply_all_filters()

    # --------------------------------------------------------
    #                  ЗАСТОСУВАННЯ ФІЛЬТРІВ
    # --------------------------------------------------------

    def apply_all_filters(self):
        if self.df_original is None:
            return

        df = self.df_original.copy()

        # 1) умови
        if self.conditions:
            df = apply_filters(df, self.conditions)

        # 2) прокуратура
        pros = self.cb_prosecutor.currentText()
        if pros and pros != "Усі прокуратури" and "Прокуратура" in df.columns:
            df = df[df["Прокуратура"] == pros]

        # 3) глобальний пошук
        if self.global_search_text:
            text = self.global_search_text
            mask = df.apply(
                lambda col: col.astype(str).str.contains(text, case=False, na=False),
                axis=0
            ).any(axis=1)
            df = df[mask]

        # 4) показати лише "спливаючі" строки (якщо ввімкнено)
        if self.show_only_expiring and self.highlight_indices:
            df = df[df.index.isin(self.highlight_indices)]

        self.df_current = df

        model = self.table_view.model()
        if isinstance(model, PandasTableModel):
            model.update_df(self.df_current, highlight_indices=self.highlight_indices)
        else:
            self.table_view.setModel(
                PandasTableModel(
                    self.df_current,
                    edit_callback=self.on_cell_edited,
                    highlight_indices=self.highlight_indices,
                )
            )

    # --------------------------------------------------------
    #            СИНХРОНІЗАЦІЯ ПРАВОК У ТАБЛИЦІ
    # --------------------------------------------------------

    def on_cell_edited(self, orig_index, column_name: str, new_value):
        if self.df_original is None:
            return
        if orig_index in self.df_original.index and column_name in self.df_original.columns:
            self.df_original.at[orig_index, column_name] = new_value
        # При изменении дат имеет смысл пересчитать "спливаючі"
        self.recalc_expiring_marks(show_popup=False)
        self.apply_all_filters()

    # --------------------------------------------------------
    #                     ДОДАВАННЯ РЯДКА
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

        if "ID" in self.df_original.columns:
            try:
                new_id = (self.df_original["ID"].max() or 0) + 1
            except Exception:
                new_id = len(self.df_original) + 1
        else:
            new_id = len(self.df_original) + 1

        row = {"ID": new_id}
        row.update(data)

        new_row_df = pd.DataFrame([row])
        self.df_original = pd.concat([self.df_original, new_row_df], ignore_index=True)

        self.recalc_expiring_marks(show_popup=False)
        self.apply_all_filters()

    # --------------------------------------------------------
    #                        ЕКСПОРТ
    # --------------------------------------------------------

    def _format_df_for_export(self, df: pd.DataFrame) -> pd.DataFrame:
        out = df.copy()
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

                # Альбомна орієнтація
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
#                  ТОЧКА ВХОДУ
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