import sys
import os
import json
from pathlib import Path

import pandas as pd
from docx import Document

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QFileDialog,
    QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QListWidget, QTableView,
    QMessageBox, QComboBox, QLineEdit, QDialog,
    QFormLayout, QDialogButtonBox
)
from PySide6.QtCore import Qt

from .model import PandasTableModel
from .load_test_data import load_test_df
from .filters_core import (
    FilterCondition, Operator, apply_filters
)

CONFIG_PATH = Path.home() / ".table_filter_engine.json"


class AddRowDialog(QDialog):
    """Диалог для добавления новой строки в таблицу."""

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
        form.addRow("ПІБ:", self.ed_pib)

        # Дата народження
        self.ed_dob = QLineEdit()
        self.ed_dob.setPlaceholderText("дд.мм.рррр")
        form.addRow("Дата нар.:", self.ed_dob)

        # Стаття ККУ
        self.ed_article = QLineEdit()
        form.addRow("Стаття ККУ:", self.ed_article)

        # Запобіжний захід
        self.ed_measure = QLineEdit()
        form.addRow("Запобіжний захід:", self.ed_measure)

        # Виїзд за кордон
        self.cb_went_abroad = QComboBox()
        self.cb_went_abroad.addItems(["Ні", "Так"])
        form.addRow("Є виїзд за кордон:", self.cb_went_abroad)

        # Дата виїзду
        self.ed_departure = QLineEdit()
        self.ed_departure.setPlaceholderText("дд.мм.рррр або порожньо")
        form.addRow("Дата виїзду:", self.ed_departure)

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
            "Є_виїзд_за_кордон": True if self.cb_went_abroad.currentText() == "Так" else False,
            "Дата_виїзду": self.ed_departure.text().strip(),
            "Країна_виїзду": self.ed_country.text().strip(),
            "Є_Інтерпол": True if self.cb_interpol.currentText() == "Так" else False,
            "Дата_оголошення_в_розшук": self.ed_interpol_date.text().strip(),
            "Примітка": self.ed_note.text().strip(),
        }


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Table Filter Engine — Prototype")
        self.resize(1400, 850)

        self.df_original: pd.DataFrame | None = None
        self.df_current: pd.DataFrame | None = None

        self.conditions: list[FilterCondition] = []
        self.global_search_text: str = ""

        self._init_ui()
        self._load_last_file_if_any()

    # ----------------- UI -----------------

    def _init_ui(self):
        central = QWidget(self)
        root_v = QVBoxLayout(central)
        root_v.setContentsMargins(5, 5, 5, 5)

        # ----- верхняя панель -----
        top_bar = QHBoxLayout()

        self.btn_load = QPushButton("📂 Відкрити таблицю")
        self.btn_load.clicked.connect(self.choose_and_load_table)
        top_bar.addWidget(self.btn_load)

        self.btn_add_row = QPushButton("➕ Додати рядок")
        self.btn_add_row.clicked.connect(self.add_row)
        self.btn_add_row.setEnabled(False)
        top_bar.addWidget(self.btn_add_row)

        self.btn_export = QPushButton("💾 Експорт")
        self.btn_export.clicked.connect(self.export_current)
        self.btn_export.setEnabled(False)
        top_bar.addWidget(self.btn_export)

        top_bar.addStretch(1)

        lbl_search = QLabel("Глобальний пошук:")
        top_bar.addWidget(lbl_search)

        self.ed_global_search = QLineEdit()
        self.ed_global_search.setPlaceholderText("Пошук по всіх стовпцях...")
        self.ed_global_search.textChanged.connect(self.on_global_search_changed)
        self.ed_global_search.setEnabled(False)
        top_bar.addWidget(self.ed_global_search, stretch=2)

        root_v.addLayout(top_bar)

        # ----- основная часть -----
        main_h = QHBoxLayout()

        # левая панель
        left = QVBoxLayout()
        left.setAlignment(Qt.AlignTop)

        # Фильтр по прокуратуре
        lbl_pros = QLabel("Фільтр по прокуратурі:")
        lbl_pros.setStyleSheet("font-weight: bold;")
        left.addWidget(lbl_pros)

        self.cb_prosecutor = QComboBox()
        self.cb_prosecutor.addItem("Усі прокуратури")
        self.cb_prosecutor.currentIndexChanged.connect(self.apply_all_filters)
        self.cb_prosecutor.setEnabled(False)
        left.addWidget(self.cb_prosecutor)

        left.addSpacing(15)

        # Конструктор фильтра по столбцу
        lbl_col_filter = QLabel("Фільтр по стовпцю:")
        lbl_col_filter.setStyleSheet("font-weight: bold;")
        left.addWidget(lbl_col_filter)

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

        # Выпадающий список возможных значений (например, для Стаття_ККУ)
        self.cb_value_choices = QComboBox()
        self.cb_value_choices.setVisible(False)
        self.cb_value_choices.currentIndexChanged.connect(self.on_value_choice_selected)
        left.addWidget(self.cb_value_choices)

        # Поля дат (для стовпців з датами)
        self.ed_date_from = QLineEdit()
        self.ed_date_from.setPlaceholderText("з дд.мм.рррр")
        self.ed_date_from.setVisible(False)
        left.addWidget(self.ed_date_from)

        self.ed_date_to = QLineEdit()
        self.ed_date_to.setPlaceholderText("по дд.мм.рррр")
        self.ed_date_to.setVisible(False)
        left.addWidget(self.ed_date_to)

        self.btn_add_condition = QPushButton("Додати умову")
        self.btn_add_condition.clicked.connect(self.add_condition_from_ui)
        self.btn_add_condition.setEnabled(False)
        left.addWidget(self.btn_add_condition)

        left.addSpacing(10)

        # Список активных условий
        lbl_current = QLabel("Поточні умови:")
        left.addWidget(lbl_current)

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

        # двойной клик по условию = удалить
        self.list_conditions.itemDoubleClicked.connect(
            lambda _: self.remove_selected_condition()
        )

        # Таблица справа
        self.table_view = QTableView()
        self.table_view.setAlternatingRowColors(True)
        self.table_view.horizontalHeader().setStretchLastSection(True)

        main_h.addLayout(left, 1)
        main_h.addWidget(self.table_view, 4)

        root_v.addLayout(main_h)
        self.setCentralWidget(central)

    # ----------------- конфиг -----------------

    def _load_last_file_if_any(self):
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

    # ----------------- загрузка таблицы -----------------

    def choose_and_load_table(self):
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

            model = PandasTableModel(self.df_current, edit_callback=self.on_cell_edited)
            self.table_view.setModel(model)

            # прокуратури
            self.cb_prosecutor.setEnabled(True)
            self.cb_prosecutor.clear()
            self.cb_prosecutor.addItem("Усі прокуратури")
            if "Прокуратура" in df.columns:
                for p in sorted(df["Прокуратура"].dropna().unique()):
                    self.cb_prosecutor.addItem(str(p))

            # стовпці
            self.cb_column.setEnabled(True)
            self.cb_column.clear()
            for col in df.columns:
                self.cb_column.addItem(col)

            self.cb_operator.setEnabled(True)
            self.ed_value.setEnabled(True)
            self.btn_add_condition.setEnabled(True)
            self.btn_clear_conditions.setEnabled(True)
            self.btn_remove_condition.setEnabled(True)

            self.btn_add_row.setEnabled(True)
            self.btn_export.setEnabled(True)
            self.ed_global_search.setEnabled(True)

            self.conditions.clear()
            self.list_conditions.clear()
            self.global_search_text = ""
            self.ed_global_search.clear()

            # подстроить режим ввода под первый столбец
            self.on_column_changed(self.cb_column.currentIndex())

            self._save_last_file(path)

            if show_message:
                QMessageBox.information(self, "OK", f"Файл завантажено:\n{path}")

        except Exception as e:
            QMessageBox.critical(self, "Помилка завантаження", str(e))

    # ----------------- глобальний пошук -----------------

    def on_global_search_changed(self, text: str):
        self.global_search_text = text.strip()
        self.apply_all_filters()

    # ----------------- переключение режима ввода по стовпцю -----------------

    def on_column_changed(self, index: int):
        if self.df_original is None or index < 0:
            return

        column = self.cb_column.itemText(index)
        series = self.df_original[column]

        is_date = pd.api.types.is_datetime64_any_dtype(series)

        if is_date:
            # режим дат
            self.cb_operator.setVisible(False)
            self.ed_value.setVisible(False)
            self.cb_value_choices.setVisible(False)
            self.ed_date_from.setVisible(True)
            self.ed_date_to.setVisible(True)
        else:
            self.cb_operator.setVisible(True)
            self.ed_value.setVisible(True)
            self.ed_date_from.setVisible(False)
            self.ed_date_to.setVisible(False)

            uniques = series.dropna().unique()
            if len(uniques) <= 50 or column == "Стаття_ККУ":
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

    # ----------------- конструктор умов -----------------

    def add_condition_from_ui(self):
        if self.df_original is None:
            return

        column = self.cb_column.currentText()
        if not column:
            return

        series = self.df_original[column]

        # столбец-дата → диапазон
        if pd.api.types.is_datetime64_any_dtype(series):
            from_text = self.ed_date_from.text().strip()
            to_text = self.ed_date_to.text().strip()

            if not from_text and not to_text:
                return

            def parse_date(txt: str):
                if not txt:
                    return None
                try:
                    return pd.to_datetime(txt, format="%d.%m.%Y")
                except Exception:
                    QMessageBox.warning(
                        self,
                        "Невірний формат дати",
                        "Використовуйте формат дд.мм.рррр",
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
            self.list_conditions.addItem(
                f"{column}: {from_text or '...'} — {to_text or '...'}"
            )
            self.apply_all_filters()
            return

        # обычные (не-даты)
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

        value: object = raw_value

        try:
            if pd.api.types.is_bool_dtype(series):
                v = raw_value.lower()
                if v in ("так", "true", "1"):
                    value = True
                elif v in ("ні", "false", "0", "нет", "no"):
                    value = False
            elif pd.api.types.is_datetime64_any_dtype(series):
                value = pd.to_datetime(raw_value, format="%d.%m.%Y")
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

    # ----------------- применение всех фильтров -----------------

    def apply_all_filters(self):
        if self.df_original is None:
            return

        df = self.df_original.copy()

        # 1) условия конструктора
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
                axis=0
            ).any(axis=1)
            df = df[mask]

        self.df_current = df

        model = self.table_view.model()
        if isinstance(model, PandasTableModel):
            model.update_df(self.df_current)
        else:
            self.table_view.setModel(
                PandasTableModel(self.df_current, edit_callback=self.on_cell_edited)
            )

    # ----------------- синхронизация правок -----------------

    def on_cell_edited(self, orig_index, column_name: str, new_value):
        """Вызывается моделью, когда пользователь меняет ячейку."""
        if self.df_original is None:
            return
        if orig_index in self.df_original.index and column_name in self.df_original.columns:
            self.df_original.at[orig_index, column_name] = new_value

    # ----------------- добавление строки -----------------

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
            new_id = (self.df_original["ID"].max() or 0) + 1
        else:
            new_id = len(self.df_original) + 1

        row = {"ID": new_id, **data}
        new_row_df = pd.DataFrame([row])
        self.df_original = pd.concat([self.df_original, new_row_df], ignore_index=True)

        self.apply_all_filters()

    # ----------------- экспорт -----------------

    def _format_df_for_export(self, df: pd.DataFrame) -> pd.DataFrame:
        """Подготовка данных к экспорту (даты и bool → человекочитаемо)."""
        out = df.copy()
        for col in out.columns:
            if pd.api.types.is_datetime64_any_dtype(out[col]):
                out[col] = out[col].dt.strftime("%d.%m.%Y").fillna("")
            elif pd.api.types.is_bool_dtype(out[col]):
                out[col] = out[col].map({True: "Так", False: "Ні"})
        return out

    def export_current(self):
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
                table = doc.add_table(rows=1, cols=len(df_out.columns))
                table.style = "Table Grid"  # чёткие границы

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