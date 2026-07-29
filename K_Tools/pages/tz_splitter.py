import os
import sys


# PyInstaller keeps pyogrio's GDAL dependencies in a nested directory.  Unlike
# a regular wheel installation, that directory is not always on Windows' DLL
# search path in a one-file build.  Keep the returned handles alive for the
# lifetime of the process; closing them removes the directory again.
_DLL_DIRECTORY_HANDLES = []
if sys.platform == "win32" and getattr(sys, "frozen", False):
    bundle_dir = getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
    pyogrio_libs = os.path.join(bundle_dir, "pyogrio.libs")
    if os.path.isdir(pyogrio_libs):
        _DLL_DIRECTORY_HANDLES.append(os.add_dll_directory(pyogrio_libs))
        os.environ["PATH"] = pyogrio_libs + os.pathsep + os.environ.get("PATH", "")
    gdal_data = os.path.join(bundle_dir, "pyogrio", "gdal_data")
    proj_data = os.path.join(bundle_dir, "pyogrio", "proj_data")
    if os.path.isdir(gdal_data):
        os.environ["GDAL_DATA"] = gdal_data
    if os.path.isdir(proj_data):
        os.environ["PROJ_DATA"] = proj_data
        os.environ["PROJ_LIB"] = proj_data

import geopandas as gpd
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QComboBox,
    QFileDialog,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QSlider,
)

from core import BasePage, PathEdit


class TzSplitterPage(BasePage):
    TAB_BOUNDS = "-1000000,-1000000,19000000,19000000"

    def __init__(self, controller=None, parent=None):
        super().__init__(controller, parent)
        root = self.page_layout(
            "Разделение территориальных зон",
            "Распределяет объекты ТЗ по населённым пунктам на основе доли пересечения.",
        )
        _, inputs = self.card_layout(root, "Входные данные")
        self.np_edit = self._file_row(inputs, "TAB с границами НП", self.select_np_file)
        self.tz_edit = self._file_row(inputs, "TAB с территориальными зонами", self.select_tz_file)
        self.output_edit = self._file_row(inputs, "Папка результата", self.select_output_folder)
        inputs.addWidget(QLabel("Поле с названием населённого пункта"))
        self.field_combo = QComboBox()
        inputs.addWidget(self.field_combo)

        _, thresholds = self.card_layout(root, "Пороги классификации")
        grid = QGridLayout()
        self.in_slider, self.in_value = self._slider(grid, 0, "В НП ≥", 50, 100, 98)
        self.out_slider, self.out_value = self._slider(grid, 1, "Вне НП <", 0, 50, 2)
        thresholds.addLayout(grid)
        reset = QPushButton("Сбросить пороги")
        reset.clicked.connect(self.reset_thresholds)
        thresholds.addWidget(reset)
        note = QLabel(
            "Между двумя порогами объект попадёт в файл «СОМНИТЕЛЬНЫЕ» для ручной проверки."
        )
        note.setWordWrap(True)
        thresholds.addWidget(note)

        self.run_btn = QPushButton("Запустить разделение")
        self.run_btn.setProperty("primary", True)
        self.run_btn.clicked.connect(self.start_processing)
        root.addWidget(self.run_btn)
        root.addWidget(self.setup_progress_bar())
        _, logs = self.card_layout(root, "Журнал")
        logs.addWidget(self.setup_log_area())

    def _file_row(self, layout, label, callback):
        layout.addWidget(QLabel(label))
        row = QHBoxLayout()
        edit = PathEdit()
        button = QPushButton("Выбрать")
        button.clicked.connect(callback)
        row.addWidget(edit, 1)
        row.addWidget(button)
        layout.addLayout(row)
        return edit

    @staticmethod
    def _slider(layout, row, title, minimum, maximum, value):
        label, result = QLabel(title), QLabel(f"{value / 100:.2f}")
        slider = QSlider(Qt.Orientation.Horizontal)
        slider.setRange(minimum, maximum)
        slider.setValue(value)
        slider.valueChanged.connect(lambda current: result.setText(f"{current / 100:.2f}"))
        layout.addWidget(label, row, 0)
        layout.addWidget(slider, row, 1)
        layout.addWidget(result, row, 2)
        return slider, result

    def select_np_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Границы НП", "", "MapInfo TAB (*.tab)")
        if path:
            self.np_edit.setText(path)
            self.load_np_fields(path)

    def select_tz_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Территориальные зоны", "", "MapInfo TAB (*.tab)")
        if path:
            self.tz_edit.setText(path)

    def select_output_folder(self):
        path = QFileDialog.getExistingDirectory(self, "Папка результата")
        if path:
            self.output_edit.setText(path)

    def load_np_fields(self, path):
        try:
            fields = [column for column in gpd.read_file(path).columns if column != "geometry"]
            self.field_combo.clear()
            self.field_combo.addItems(fields)
            self.log(f"Поля НП: {fields}")
        except Exception as error:
            self.log(f"Ошибка чтения полей: {error}")

    def reset_thresholds(self):
        self.in_slider.setValue(98)
        self.out_slider.setValue(2)

    @staticmethod
    def to_safe_filename(name):
        for char in r'/\\:*?"<>|':
            name = name.replace(char, "_")
        return name.strip()

    @staticmethod
    def fix_geom(frame):
        frame = frame.copy()
        frame["geometry"] = frame.geometry.buffer(0)
        return frame[~frame.geometry.is_empty & frame.is_valid].reset_index(drop=True)

    @classmethod
    def save_tab(cls, frame, path):
        frame.to_file(
            path,
            driver="MapInfo File",
            engine="pyogrio",
            layer_options={"BOUNDS": cls.TAB_BOUNDS},
        )

    def start_processing(self):
        params = (
            self.np_edit.text().strip(),
            self.tz_edit.text().strip(),
            self.output_edit.text().strip(),
            self.field_combo.currentText().strip(),
            self.in_slider.value() / 100,
            self.out_slider.value() / 100,
        )
        if not all(params[:4]):
            QMessageBox.critical(self, "Ошибка", "Не все поля заполнены.")
            return
        if not os.path.isfile(params[0]) or not os.path.isfile(params[1]):
            QMessageBox.critical(self, "Ошибка", "Один из TAB-файлов не существует.")
            return
        self.clear_log()
        self.run_btn.setEnabled(False)
        self.start_task(
            self._process,
            *params,
            on_result=self._done,
            on_error=self._error,
            on_finished=lambda: self.run_btn.setEnabled(True),
        )

    @classmethod
    def _process(cls, signals, np_path, tz_path, out_dir, name_field, threshold_in, threshold_out):
        os.makedirs(out_dir, exist_ok=True)
        signals.message.emit(
            f"Пороги: в НП ≥ {threshold_in:.0%}, вне НП < {threshold_out:.0%}"
        )
        signals.message.emit("Загрузка данных…")
        np_gdf = cls.fix_geom(gpd.read_file(np_path))
        tz_gdf = cls.fix_geom(gpd.read_file(tz_path))
        signals.message.emit(f"НП: {len(np_gdf)}, ТЗ: {len(tz_gdf)}")
        if name_field not in np_gdf.columns:
            raise ValueError(f"Поле «{name_field}» отсутствует в таблице НП")
        if np_gdf.crs != tz_gdf.crs:
            signals.message.emit("Преобразование CRS…")
            tz_gdf = tz_gdf.to_crs(np_gdf.crs)

        np_results = {name: [] for name in np_gdf[name_field].dropna().unique()}
        outside, questionable = [], []
        for position, (index, tz_row) in enumerate(tz_gdf.iterrows(), 1):
            geometry, area = tz_row.geometry, tz_row.geometry.area
            intersections = []
            for _, np_row in np_gdf.iterrows():
                intersection = geometry.intersection(np_row.geometry)
                if not intersection.is_empty:
                    intersections.append(
                        (np_row[name_field], intersection.area / area if area > 0 else 0)
                    )
            intersections.sort(key=lambda item: item[1], reverse=True)
            if not intersections:
                outside.append(tz_row)
            else:
                best_np, best_share = intersections[0]
                if best_share >= threshold_in:
                    np_results.setdefault(best_np, []).append(tz_row)
                elif best_share < threshold_out:
                    outside.append(tz_row)
                else:
                    row_copy = tz_row.copy()
                    reason = f"{best_share:.1%} в НП «{best_np}»"
                    if len(intersections) > 1 and intersections[1][1] > threshold_out:
                        reason += f", {intersections[1][1]:.1%} в НП «{intersections[1][0]}»"
                    row_copy["ПРИЧИНА"] = reason
                    questionable.append(row_copy)
                    signals.message.emit(f"ТЗ {index}: {reason} → СОМНИТЕЛЬНЫЕ")
            signals.progress.emit(position, len(tz_gdf))

        saved_np = 0
        for name, rows in np_results.items():
            if rows:
                output = gpd.GeoDataFrame(rows, crs=tz_gdf.crs)
                cls.save_tab(
                    output,
                    os.path.join(out_dir, f"{cls.to_safe_filename(str(name))}.tab"),
                )
                saved_np += 1
                signals.message.emit(f"НП «{name}»: сохранено {len(rows)}")
        if outside:
            cls.save_tab(
                gpd.GeoDataFrame(outside, crs=tz_gdf.crs),
                os.path.join(out_dir, "ВНЕ_НП.tab"),
            )
        if questionable:
            cls.save_tab(
                gpd.GeoDataFrame(questionable, crs=tz_gdf.crs),
                os.path.join(out_dir, "СОМНИТЕЛЬНЫЕ.tab"),
            )
        return saved_np, len(outside), len(questionable)

    def _done(self, result):
        np_count, outside, questionable = result
        self.log("Обработка завершена.")
        QMessageBox.information(
            self,
            "Готово",
            f"Файлов НП: {np_count}\nВне НП: {outside} объектов\n"
            f"Сомнительные: {questionable} объектов",
        )

    def _error(self, text):
        self.log(text)
        QMessageBox.critical(self, "Ошибка", text.splitlines()[-1])
