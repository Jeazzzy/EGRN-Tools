import os
import re
import shutil

from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from core import BasePage, IndexSelector, PathEdit

try:
    import pyodbc
    PYODBC_AVAILABLE = True
except ImportError:
    pyodbc = None
    PYODBC_AVAILABLE = False


class MdbCopyPage(BasePage):
    TYPE_GENITIVE = {
        "аал": "аала", "автодорога": "автодороги", "арбан": "арбана",
        "аул": "аула", "волость": "волости", "высел": "выселок",
        "г": "города", "городок": "городка", "д": "деревни",
        "дп": "дачного поселка", "ж/д_будка": "железнодорожной будки",
        "ж/д_казарм": "железнодорожной казармы",
        "ж/д_оп": "ж/д остановочного (обгонного) пункта",
        "ж/д_платф": "железнодорожной платформы",
        "ж/д_пост": "железнодорожного поста",
        "ж/д_рзд": "железнодорожного разъезда",
        "ж/д_ст": "железнодорожной станции", "жилзона": "жилой зоны",
        "жилрайон": "жилого района", "заимка": "заимки", "казарма": "казармы",
        "кв-л": "квартала", "кордон": "кордона", "кп": "курортного поселка",
        "лпх": "леспромхоза", "м": "местечка", "массив": "массива",
        "мкр": "микрорайона", "нп": "населенного пункта", "остров": "острова",
        "п": "поселка", "п/о": "почтового отделения",
        "п/р": "планировочного района", "п/ст": "поселка и(при) станции(и)",
        "пгт": "поселка городского типа", "погост": "погоста",
        "починок": "починка", "промзона": "промышленной зоны",
        "рзд": "разъезда", "рп": "рабочего поселка", "с": "села",
        "сл": "слободы", "снт": "садового некоммерческого товарищества",
        "ст": "станции", "ст-ца": "станицы", "тер": "территории",
        "у": "улуса", "х": "хутора",
    }

    def __init__(self, controller=None, parent=None):
        super().__init__(controller, parent)
        root = self.page_layout(
            "Работа с MDB",
            "Массовая замена файлов и таблиц Microsoft Access.",
        )
        if not PYODBC_AVAILABLE:
            warning = QLabel("pyodbc не установлен. Добавьте зависимость перед работой с MDB.")
            warning.setStyleSheet(
                "background:#fff3cd;color:#856404;padding:12px;border-radius:8px;"
            )
            root.addWidget(warning)

        self.tabs = QTabWidget()
        root.addWidget(self.tabs)
        self._build_replace_mdb()
        self._build_vri()
        self._build_replace_table()
        self._build_fias()
        self._build_text()
        self.run_btn = QPushButton("Запустить операцию")
        self.run_btn.setProperty("primary", True)
        self.run_btn.clicked.connect(self.run_operation)
        root.addWidget(self.run_btn)
        root.addWidget(self.setup_progress_bar())
        _, logs = self.card_layout(root, "Журнал")
        logs.addWidget(self.setup_log_area())

    @staticmethod
    def _tab():
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(16, 16, 16, 16)
        return widget, layout

    def _directory_row(self, layout, label, source_selector=None):
        layout.addWidget(QLabel(label))
        row = QHBoxLayout()
        edit = PathEdit()
        button = QPushButton("Выбрать")
        if source_selector is None:
            button.clicked.connect(lambda: self._pick_dir(edit))
        else:
            button.clicked.connect(lambda: self._pick_source_dir(edit, source_selector))
        row.addWidget(edit, 1)
        row.addWidget(button)
        layout.addLayout(row)
        return edit

    def _file_row(self, layout, label, callback):
        layout.addWidget(QLabel(label))
        row = QHBoxLayout()
        edit = PathEdit()
        button = QPushButton("Выбрать MDB")
        button.clicked.connect(lambda: callback(edit))
        row.addWidget(edit, 1)
        row.addWidget(button)
        layout.addLayout(row)
        return edit

    def _build_replace_mdb(self):
        tab, layout = self._tab()
        note = QLabel("Полностью заменяет target MDB файлом из папки соответствующего индекса.")
        note.setWordWrap(True)
        layout.addWidget(note)
        self.replace_selector = IndexSelector()
        self.replace_source = self._directory_row(
            layout, "Source: папки-индексы с MDB", self.replace_selector
        )
        layout.addWidget(QLabel("Индексы для обработки"))
        layout.addWidget(self.replace_selector)
        self.replace_target = self._directory_row(
            layout, "Target: структура МО / НП / индекс / … / MDB"
        )
        self.tabs.addTab(tab, "Замена MDB")

    def _build_vri(self):
        tab, layout = self._tab()
        note = QLabel("Копирует таблицу Utilizations_KP между MDB одинакового индекса.")
        layout.addWidget(note)
        self.vri_selector = IndexSelector()
        self.vri_source = self._directory_row(
            layout, "Source: папки-индексы с MDB", self.vri_selector
        )
        layout.addWidget(QLabel("Индексы для обработки"))
        layout.addWidget(self.vri_selector)
        self.vri_target = self._directory_row(layout, "Target: папка с MDB")
        self.tabs.addTab(tab, "ВРИ")

    def _build_replace_table(self):
        tab, layout = self._tab()
        self.table_source = self._file_row(
            layout, "Source MDB", self._pick_table_source
        )
        self.table_target = self._directory_row(layout, "Target: папка с MDB")
        layout.addWidget(QLabel("Имя таблицы"))
        row = QHBoxLayout()
        self.table_combo = QComboBox()
        self.table_combo.setEditable(True)
        refresh = QPushButton("Обновить список")
        refresh.clicked.connect(self._load_tables)
        row.addWidget(self.table_combo, 1)
        row.addWidget(refresh)
        layout.addLayout(row)
        self.tabs.addTab(tab, "Одна таблица")

    def _build_fias(self):
        tab, layout = self._tab()
        layout.addWidget(QLabel("Обновляет таблицу Locations во всех target MDB."))
        self.fias_source = self._file_row(layout, "Source MDB", self._pick_mdb)
        self.fias_target = self._directory_row(layout, "Target: папка с MDB")
        self.tabs.addTab(tab, "Адрес по ФИАС")

    def _build_text(self):
        tab, layout = self._tab()
        layout.addWidget(
            QLabel("Обновляет поле Объект_ЗУ таблицы Титульный_картаплан по Locations.")
        )
        self.text_folder = self._directory_row(layout, "Папка с MDB")
        self.text_outside = QCheckBox("Вне НП — текст без названия населённого пункта")
        layout.addWidget(self.text_outside)
        self.tabs.addTab(tab, "Адрес в тексте")

    def _pick_dir(self, edit):
        path = QFileDialog.getExistingDirectory(self, "Выберите папку")
        if path:
            edit.setText(path)

    def _pick_source_dir(self, edit, selector):
        self._pick_dir(edit)
        root = edit.text().strip()
        if os.path.isdir(root):
            selector.load(sorted(
                name for name in os.listdir(root)
                if os.path.isdir(os.path.join(root, name))
            ))

    def _pick_mdb(self, edit):
        path, _ = QFileDialog.getOpenFileName(self, "Выберите MDB", "", "MDB (*.mdb)")
        if path:
            edit.setText(path)

    def _pick_table_source(self, edit):
        self._pick_mdb(edit)
        if edit.text():
            self._load_tables()

    def _load_tables(self):
        path = self.table_source.text().strip()
        if not PYODBC_AVAILABLE:
            QMessageBox.critical(self, "Ошибка", "pyodbc не установлен.")
            return
        if not os.path.isfile(path):
            QMessageBox.warning(self, "Внимание", "Сначала выберите MDB.")
            return
        try:
            connection = self._get_conn(path)
            tables = [row.table_name for row in connection.cursor().tables(tableType="TABLE")]
            connection.close()
            self.table_combo.clear()
            self.table_combo.addItems(tables)
            self.log(f"Загружено таблиц: {len(tables)}")
        except Exception as error:
            QMessageBox.critical(self, "Ошибка", str(error))

    @staticmethod
    def _get_conn(path):
        return pyodbc.connect(
            rf"DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={path};",
            autocommit=False,
        )

    @staticmethod
    def _collect_source_by_index(root, selected):
        result = {}
        for name in os.listdir(root):
            folder = os.path.join(root, name)
            if name not in selected or not os.path.isdir(folder):
                continue
            source = next(
                (os.path.join(folder, item) for item in os.listdir(folder)
                 if item.lower().endswith(".mdb")),
                None,
            )
            if source:
                result[name] = source
        return result

    @staticmethod
    def _find_target_mdb_by_index(root, index_name):
        result = []
        for folder, _, files in os.walk(root):
            if index_name in os.path.normpath(folder).split(os.sep):
                result.extend(
                    os.path.join(folder, name)
                    for name in files if name.lower().endswith(".mdb")
                )
        return result

    @staticmethod
    def _collect_all_mdb(root):
        return [
            os.path.join(folder, name)
            for folder, _, files in os.walk(root)
            for name in files if name.lower().endswith(".mdb")
        ]

    def _copy_table(self, source, target, table):
        source_conn, target_conn = self._get_conn(source), self._get_conn(target)
        try:
            source_cursor, target_cursor = source_conn.cursor(), target_conn.cursor()
            source_cursor.execute(f"SELECT * FROM [{table}]")
            rows = source_cursor.fetchall()
            columns = [description[0] for description in source_cursor.description]
            target_cursor.execute(f"DELETE FROM [{table}]")
            target_conn.commit()
            if rows:
                names = ", ".join(f"[{name}]" for name in columns)
                placeholders = ", ".join("?" for _ in columns)
                target_cursor.executemany(
                    f"INSERT INTO [{table}] ({names}) VALUES ({placeholders})", rows
                )
                target_conn.commit()
        finally:
            source_conn.close()
            target_conn.close()

    def _get_locality(self, path):
        connection = self._get_conn(path)
        try:
            cursor = connection.cursor()
            cursor.execute("SELECT * FROM [Locations]")
            columns = [description[0] for description in cursor.description]
            row = cursor.fetchone()
            if row is None:
                return None, None
            values = dict(zip(columns, row))
            name = values.get("City_Name") or values.get("city_name")
            kind = values.get("City_Type") or values.get("city_type")
            if not name or not kind:
                name = values.get("Locality_Name") or values.get("locality_name")
                kind = values.get("Locality_Type") or values.get("locality_type")
            if not name or not kind:
                return None, None
            kind = self.TYPE_GENITIVE.get(str(kind).strip().lower(), str(kind).strip())
            return str(name).strip(), kind
        finally:
            connection.close()

    def _update_title(self, path, name, kind, outside=False):
        connection, updated = self._get_conn(path), 0
        try:
            cursor = connection.cursor()
            cursor.execute("SELECT * FROM [Титульный_картаплан]")
            columns = [description[0] for description in cursor.description]
            if "Объект_ЗУ" not in columns:
                return 0
            primary = columns[0]
            pattern = re.compile(
                r"в\s+границах\s+.*?муниципального(?:\s+образования)?",
                re.IGNORECASE | re.DOTALL,
            )
            replacement = (
                "в границах муниципального образования"
                if outside else f"в границах {kind} {name} муниципального образования"
            )
            for row in cursor.fetchall():
                values = dict(zip(columns, row))
                old = values.get("Объект_ЗУ")
                if isinstance(old, str):
                    new = pattern.sub(replacement, old)
                    if new != old:
                        cursor.execute(
                            f"UPDATE [Титульный_картаплан] "
                            f"SET [Объект_ЗУ]=? WHERE [{primary}]=?",
                            (new, values[primary]),
                        )
                        updated += 1
            connection.commit()
            return updated
        finally:
            connection.close()

    def run_operation(self):
        if not PYODBC_AVAILABLE:
            QMessageBox.critical(self, "Ошибка", "Установите pyodbc.")
            return
        mode = self.tabs.currentIndex()
        if mode == 0:
            params = (
                self.replace_source.text().strip(),
                self.replace_target.text().strip(),
                set(self.replace_selector.get_selected()),
            )
        elif mode == 1:
            params = (
                self.vri_source.text().strip(),
                self.vri_target.text().strip(),
                set(self.vri_selector.get_selected()),
            )
        elif mode == 2:
            params = (
                self.table_source.text().strip(),
                self.table_target.text().strip(),
                self.table_combo.currentText().strip(),
            )
        elif mode == 3:
            params = (self.fias_source.text().strip(), self.fias_target.text().strip())
        else:
            params = (self.text_folder.text().strip(), self.text_outside.isChecked())
        self.clear_log()
        self.run_btn.setEnabled(False)
        self.start_task(
            self._execute,
            mode,
            params,
            on_result=lambda result: QMessageBox.information(self, "Готово", result),
            on_error=self._show_error,
            on_finished=lambda: self.run_btn.setEnabled(True),
        )

    def _execute(self, signals, mode, params):
        if mode in (0, 1):
            source_root, target_root, selected = params
            if not os.path.isdir(source_root) or not os.path.isdir(target_root):
                raise ValueError("Укажите корректные папки SOURCE и TARGET")
            if not selected:
                raise ValueError("Выберите хотя бы один индекс")
            sources = self._collect_source_by_index(source_root, selected)
            if not sources:
                raise ValueError("SOURCE MDB не найдены по выбранным индексам")
            changed = 0
            for position, (index, source) in enumerate(sorted(sources.items()), 1):
                targets = self._find_target_mdb_by_index(target_root, index)
                signals.message.emit(f"Индекс {index}: target файлов — {len(targets)}")
                for target in targets:
                    try:
                        if mode == 0:
                            shutil.copy2(source, target)
                        else:
                            self._copy_table(source, target, "Utilizations_KP")
                        changed += 1
                        signals.message.emit(f"  OK: {target}")
                    except Exception as error:
                        signals.message.emit(f"  Ошибка: {target}: {error}")
                signals.progress.emit(position, len(sources))
            return f"Операция завершена. Обновлено: {changed}"

        if mode == 2:
            source, target_root, table = params
            if not os.path.isfile(source) or not os.path.isdir(target_root) or not table:
                raise ValueError("Проверьте source MDB, target папку и имя таблицы")
            targets = [
                item for item in self._collect_all_mdb(target_root)
                if os.path.abspath(item) != os.path.abspath(source)
            ]
            return self._copy_to_targets(signals, source, targets, table)

        if mode == 3:
            source, target_root = params
            if not os.path.isfile(source) or not os.path.isdir(target_root):
                raise ValueError("Проверьте source MDB и target папку")
            targets = [
                item for item in self._collect_all_mdb(target_root)
                if os.path.abspath(item) != os.path.abspath(source)
            ]
            return self._copy_to_targets(signals, source, targets, "Locations")

        folder, outside = params
        if not os.path.isdir(folder):
            raise ValueError("Укажите корректную папку с MDB")
        files = self._collect_all_mdb(folder)
        if not files:
            raise ValueError("MDB-файлы не найдены")
        total_updated = 0
        for position, path in enumerate(files, 1):
            try:
                if outside:
                    updated = self._update_title(path, None, None, True)
                else:
                    name, kind = self._get_locality(path)
                    if not name:
                        signals.message.emit(f"НП не найден: {path}")
                        signals.progress.emit(position, len(files))
                        continue
                    updated = self._update_title(path, name, kind)
                total_updated += updated
                signals.message.emit(f"{os.path.basename(path)}: обновлено строк {updated}")
            except Exception as error:
                signals.message.emit(f"{path}: {error}")
            signals.progress.emit(position, len(files))
        return f"Операция завершена. Обновлено строк: {total_updated}"

    def _copy_to_targets(self, signals, source, targets, table):
        if not targets:
            raise ValueError("Target MDB-файлы не найдены")
        changed = 0
        for position, target in enumerate(targets, 1):
            try:
                self._copy_table(source, target, table)
                changed += 1
                signals.message.emit(f"OK: {target}")
            except Exception as error:
                signals.message.emit(f"Ошибка: {target}: {error}")
            signals.progress.emit(position, len(targets))
        return f"Таблица {table} обновлена в {changed} MDB"

    def _show_error(self, text):
        self.log(text)
        QMessageBox.critical(self, "Ошибка", text.splitlines()[-1])
