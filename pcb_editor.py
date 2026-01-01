import os
import traceback

from PyQt5.QtGui import QDoubleValidator, QIntValidator

from drillingFileReader import drillingFileReader
from dxfFileReader import DXFReader
from pathlib import Path

from PyQt5.QtCore import Qt, QSettings
from PyQt5.QtWidgets import (QMainWindow, QAction, QStatusBar, QMessageBox, QFileDialog,
                             QDialog, QVBoxLayout, QLabel, QLineEdit, QHBoxLayout, QPushButton, QTreeWidget,
                             QTreeWidgetItem, QMenu, QComboBox)
from kompas_service import KompasService


class PCBEditor(QMainWindow):
    def __init__(self):
        super().__init__()
        self.addTracksAction = None
        self.project_name = None
        self.tree_view = None
        self.ks_service = KompasService()
        self.initUI()
        self.properties = QSettings("company", "pcb_editor")

    def initUI(self):
        self.setWindowTitle('Редактор печатных плат')
        self.setGeometry(100, 100, 500, 300)

        self.tree_view = QTreeWidget()
        self.tree_view.setHeaderLabel("")
        self.tree_view.setContextMenuPolicy(Qt.CustomContextMenu)
        self.tree_view.customContextMenuRequested.connect(self.show_context_menu)
        self.tree_view.clicked.connect(lambda : self.ks_service.select_macro(self.tree_view.currentItem().data(1, 0)))
        self.setCentralWidget(self.tree_view)

        self.create_menus()

        self.statusBar = QStatusBar()
        self.setStatusBar(self.statusBar)
        self.statusBar.showMessage("Готово")

    def show_context_menu(self, position):
        item = self.tree_view.itemAt(position)
        if not item or not item.parent():
            return

        parent = item.parent()
        if parent and parent.parent() is not None:
            return

        context_menu = QMenu(self)

        rename_action = QAction("Переименовать", self)
        delete_action = QAction("Удалить", self)
        control_program_action = QAction("Сформировать УП", self)
        trajectory_action = QAction("Сформировать траекторию", self)

        rename_action.triggered.connect(lambda: self.ks_service.rename_macro(item.text(0)))
        delete_action.triggered.connect(lambda: self.delete_macro(item))

        context_menu.addAction(rename_action)
        context_menu.addAction(delete_action)
        if item.text(0) == "Отверстия":
            control_program_action.triggered.connect(lambda: self.show_drill_menu(item.data(1, 0)))
            context_menu.addAction(control_program_action)
        if item.text(0) == "Границы":
            trajectory_action.triggered.connect(lambda: self.show_borders_trajectory_menu(item.data(1, 0)))
            context_menu.addAction(trajectory_action)
        if item.text(0) == "Дорожки":
            trajectory_action.triggered.connect(lambda: self.show_tracks_trajectory_menu(item.data(1, 0)))
            context_menu.addAction(trajectory_action)
        if item.text(0) == "Траектория дорожек":
            control_program_action.triggered.connect(lambda: self.show_borders_program_menu(item.data(1, 0)))
            context_menu.addAction(control_program_action)
        if item.text(0) == "Траектория границ":
            control_program_action.triggered.connect(lambda: self.show_borders_program_menu(item.data(1, 0)))
            context_menu.addAction(control_program_action)

        context_menu.exec_(self.tree_view.viewport().mapToGlobal(position))

    def delete_macro(self, item):
        self.ks_service.delete_macro(item.data(1, 0))
        self.build_project_tree(self.project_name)

    def create_menus(self):
        self.create_file_menu()
        self.create_edit_menus()

    def create_edit_menus(self):
        edit_menu = self.menuBar().addMenu('Редактировать')
        self.create_point_action = QAction('Создать начальную точку', self)
        self.create_point_action.triggered.connect(self.create_start_point)
        self.create_point_action.setDisabled(True)
        edit_menu.addAction(self.create_point_action)

    def create_file_menu(self):
        fileMenu = self.menuBar().addMenu('Файл')

        newAction = QAction('Создать проект...', self)
        newAction.setShortcut('Ctrl+N')
        newAction.triggered.connect(self.new_project)
        fileMenu.addAction(newAction)

        openAction = QAction('Открыть проект...', self)
        openAction.setShortcut('Ctrl+O')
        openAction.triggered.connect(self.open_project)
        fileMenu.addAction(openAction)

        fileMenu.addSeparator()

        self.addTracksAction = QAction('Добавить дорожки', self)
        self.addTracksAction.triggered.connect(self.add_tracks)
        fileMenu.addAction(self.addTracksAction)
        self.addTracksAction.setDisabled(True)

        self.addBordersAction = QAction('Добавить границы', self)
        self.addBordersAction.triggered.connect(self.add_borders)
        fileMenu.addAction(self.addBordersAction)
        self.addBordersAction.setDisabled(True)

        self.addHolesAction = QAction('Добавить отверстия', self)
        self.addHolesAction.triggered.connect(self.add_holes)
        fileMenu.addAction(self.addHolesAction)
        self.addHolesAction.setDisabled(True)
        
        self.addMasksAction = QAction('Добавить маску', self)
        self.addMasksAction.triggered.connect(self.add_mask)
        fileMenu.addAction(self.addMasksAction)
        self.addMasksAction.setDisabled(True)

        self.exitAction = QAction('Выход', self)
        self.exitAction.setShortcut('Ctrl+Q')
        self.exitAction.triggered.connect(self.close)
        fileMenu.addAction(self.exitAction)

    def open_project(self):
        print("open project")
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Открыть проект",
            "",
            "PCB Project Files (*.pcbprj)"
        )
        if file_path:
            self.project_name = os.path.splitext(os.path.basename(file_path))[0]
            self.ks_service.open_fragment(file_path.replace(".pcbprj", ".frw"))
            self.build_project_tree(self.project_name)
            
            doc = self.ks_service.doc
            doc2d = self.ks_service.kompas_api7_module.IKompasDocument2D(doc)

            views = doc2d.ViewsAndLayersManager.Views.View(0)

            container = self.ks_service.kompas_api7_module.IDrawingContainer(views)
            print(self.ks_service.kompas_api7_module.ksDrawingObjectNotify, self.create_menus, container.MacroObjects)
            self.ks_service.advise_kompas_event(self.ks_service.kompas_api7_module.ksDrawingObjectNotify, lambda a, b: self.build_project_tree(self.project_name), container.MacroObjects)

    def new_project(self):
        print("Создать проект")

        dlg = QDialog(self)
        dlg.setWindowTitle("Создать новый проект")
        dlg.setModal(True)
        dlg.setFixedSize(500, 250)

        layout = QVBoxLayout()

        name_label = QLabel("Название проекта:")
        name_input = QLineEdit()
        name_input.setPlaceholderText("Введите название проекта")
        layout.addWidget(name_label)
        layout.addWidget(name_input)

        folder_label = QLabel("Папка проекта:")
        folder_layout = QHBoxLayout()

        folder_input = QLineEdit()
        folder_input.setPlaceholderText("Выберите папку для сохранения")
        folder_button = QPushButton("Обзор...")

        folder_layout.addWidget(folder_input)
        folder_layout.addWidget(folder_button)

        layout.addWidget(folder_label)
        layout.addLayout(folder_layout)

        button_layout = QHBoxLayout()
        create_button = QPushButton("Создать")
        cancel_button = QPushButton("Отмена")

        button_layout.addWidget(create_button)
        button_layout.addWidget(cancel_button)
        layout.addLayout(button_layout)

        layout.addStretch()

        dlg.setLayout(layout)

        def browse_folder():
            folder = QFileDialog.getExistingDirectory(
                dlg,
                "Выберите папку для проекта",
                "",
                QFileDialog.ShowDirsOnly
            )
            if folder:
                folder_input.setText(folder)

        def create_project():
            self.project_name = name_input.text().strip()
            project_folder = folder_input.text().strip()

            if not self.project_name:
                QMessageBox.warning(dlg, "Ошибка", "Введите название проекта!")
                return

            if not project_folder:
                QMessageBox.warning(dlg, "Ошибка", "Выберите папку для проекта!")
                return

            project_path = Path(project_folder)

            if not os.path.exists(project_path):
                QMessageBox.warning(dlg, "Ошибка", f"Папка '{project_path}' не существует!")
                return

            try:
                project_path = project_path / f"{self.project_name}"
                if os.path.exists(project_path):
                    response = QMessageBox.question(
                        dlg,
                        "Подтверждение",
                        "Проект уже существует. Перезаписать?",
                        QMessageBox.Yes | QMessageBox.No,
                        QMessageBox.No
                    )
                    if response != QMessageBox.Yes:
                        return
                os.makedirs(project_path, exist_ok=True)
                frw_path = project_path / f"{self.project_name}.frw"
                pcbprj_path = project_path / f"{self.project_name}.pcbprj"
                print(frw_path)


                self.ks_service.create_fragment(frw_path)
                self.build_project_tree(self.project_name)
                with open(pcbprj_path, 'w'): pass

                self.statusBar.showMessage(f"Создан проект: {project_path}")
                QMessageBox.information(dlg, "Успех", f"Проект '{self.project_name}' создан успешно!")
                dlg.accept()

            except Exception as e:
                QMessageBox.critical(dlg, "Ошибка", f"Не удалось создать проект: {str(e)}")

        def cancel():
            dlg.reject()

        folder_button.clicked.connect(browse_folder)
        create_button.clicked.connect(create_project)
        cancel_button.clicked.connect(cancel)

        dlg.exec()

    def build_project_tree(self, project_name="project"):
        print("Построение дерева проекта...")
        self.tree_view.clear()

        try:
            root = QTreeWidgetItem(self.tree_view)
            root.setText(0, project_name)
            root.setExpanded(True)
            for macro in self.ks_service.get_macros():
                keeper = self.ks_service.kompas_api7_module.IPropertyKeeper(macro)
                prop = self.ks_service.property_mng.GetProperty(self.ks_service.doc, "Тип")
                value = keeper.GetPropertyValue(prop, None, True, True)
                macro_item = QTreeWidgetItem(root)
                macro_item.setText(0, value[1])
                macro_item.setData(1, 0, macro)

            self.create_point_action.setDisabled(False)
            self.set_action_enable()
            
        except Exception as e:
            self.statusBar.showMessage("Произошла ошибка при создании дерева проекта")
            print("Ошибка:", str(e))
            traceback.print_exc()

    def add_tracks(self):
        self.statusBar.showMessage("Режим добавления дорожек")
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Открыть файл дорожек",
            "",
            "Tracks Files (*.dxf)"
        )

        if not file_path:
            self.statusBar.showMessage("Файл не выбран")
            return

        try:
            tracks_data = DXFReader.readFile(file_path)
            self.ks_service.draw_tracks(tracks_data)
            self.statusBar.showMessage(f"Дорожки загружены из: {os.path.basename(file_path)}")
        except Exception as e:
            self.statusBar.showMessage(f"Ошибка загрузки: {str(e)}")
            print(f"Ошибка при чтении файла дорожек: {e}")
    
    def add_mask(self):
        self.statusBar.showMessage("Режим добавления границ")
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Открыть файл границ",
            "",
            "Mask Files (*.dxf)"
        )

        if not file_path:
            self.statusBar.showMessage("Файл не выбран")
            return

        try:
            mask_data = DXFReader.readFile(file_path)
            self.ks_service.draw_mask(mask_data)
            self.statusBar.showMessage(f"Маска загружена из: {os.path.basename(file_path)}")
        except Exception as e:
            self.statusBar.showMessage(f"Ошибка загрузки: {str(e)}")
            print(f"Ошибка при чтении файла маски: {e}")
    
    def add_borders(self):
        self.statusBar.showMessage("Режим добавления границ")
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Открыть файл границ",
            "",
            "Borders Files (*.dxf)"
        )

        if not file_path:
            self.statusBar.showMessage("Файл не выбран")
            return

        try:
            border_data = DXFReader.readFile(file_path)
            self.ks_service.draw_border(border_data)
            self.statusBar.showMessage(f"Границы загружены из: {os.path.basename(file_path)}")
        except Exception as e:
            self.statusBar.showMessage(f"Ошибка загрузки: {str(e)}")
            print(f"Ошибка при чтении файла границ: {e}")

    def open_holes(self):
        print("open holes")
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Открыть файл сверловки",
            "",
            "Holes Files (*.drl)"
        )
        return file_path

    def add_holes(self):
        path = self.open_holes()
        if path:
            try:
                holes = drillingFileReader.readFile(path)
                self.ks_service.create_holes(holes)
                self.statusBar.showMessage("Режим добавления отверстий")
            except Exception as e:
                self.statusBar.showMessage("Произошла ошибка при добавлении отверстий")
                print("Ошибка:", str(e))
                traceback.print_exc()
        self.build_project_tree(self.project_name)

    def show_drill_menu(self, macro):
        """Метод для ввода параметров отверстия"""
        dialog = QDialog(self)
        dialog.setWindowTitle("Меню параметров")
        dialog.setModal(True)
        dialog.setFixedSize(300, 180)

        layout = QVBoxLayout()

        fields = []
        labels = ["Глубина:", "Высота перебега:", "Скорость подачи:"]

        validator = QDoubleValidator()
        validator.setNotation(QDoubleValidator.StandardNotation)

        for label_text in labels:
            row = QHBoxLayout()
            label = QLabel(label_text)
            line_edit = QLineEdit()
            line_edit.setPlaceholderText("0.0 мм")
            self.set_default_value(line_edit, label_text)

            line_edit.setValidator(validator)

            row.addWidget(label)
            row.addWidget(line_edit)
            layout.addLayout(row)
            fields.append(line_edit)

        button_layout = QHBoxLayout()
        ok_button = QPushButton("OK")
        cancel_button = QPushButton("Отмена")

        button_layout.addStretch()
        button_layout.addWidget(ok_button)
        button_layout.addWidget(cancel_button)

        layout.addLayout(button_layout)
        dialog.setLayout(layout)

        def on_ok():
            try:
                values = []
                for field in fields:
                    text = field.text().strip()
                    if not text:
                        raise ValueError("Поле не может быть пустым")
                    try:
                        val = float(text.replace(',', '.'))
                        values.append(val)
                    except ValueError:
                        raise ValueError("Некорректное число")

                depth, overrun, feedrate = values
                self.properties.setValue("depth", depth)
                self.properties.setValue("overrun", overrun)
                self.properties.setValue("feedrate", feedrate)
                print(f"Введены параметры: depth={depth}, overrun={overrun}, feedrate={feedrate}")
                program = self.ks_service.create_drilling_program(self.ks_service.find_macro_by_type("Ноль станка"), macro, depth, overrun, feedrate)
                
                file_path, _ = QFileDialog.getSaveFileName(
                    self,
                    "Сохранить файл сверловки",
                    "",
                    "CNC program (*.nc)"
                )
                
                with open(file_path, "w") as file:
                    file.write(program)
                
                dialog.accept()

            except ValueError as e:
                QMessageBox.warning(dialog, "Ошибка ввода", f"Введите корректные числа\n{e}")

        ok_button.clicked.connect(on_ok)
        cancel_button.clicked.connect(dialog.reject)

        dialog.exec_()

    def set_default_value(self, line_edit, label_text):
        key_map = {
            "Глубина": "depth",
            "Высота перебега": "overrun",
            "Скорость подачи": "feedrate"
        }
        key = key_map.get(label_text.rstrip(':'))
        print(key)
        value = self.properties.value(key)
        if value is not None:
            line_edit.setText(str(value))

    def show_borders_program_menu(self, macro):
        """Метод для ввода параметров фрезеровки границ"""
        dialog = QDialog(self)
        dialog.setWindowTitle("Меню параметров")
        dialog.setModal(True)
        dialog.setFixedSize(300, 180)

        layout = QVBoxLayout()

        fields = []
        labels = ["Глубина:", "Высота перебега:", "Скорость подачи:"]

        validator = QDoubleValidator()
        validator.setNotation(QDoubleValidator.StandardNotation)

        for label_text in labels:
            row = QHBoxLayout()
            label = QLabel(label_text)
            line_edit = QLineEdit()
            line_edit.setPlaceholderText("0.0 мм")

            line_edit.setValidator(validator)

            row.addWidget(label)
            row.addWidget(line_edit)
            layout.addLayout(row)
            fields.append(line_edit)

        button_layout = QHBoxLayout()
        ok_button = QPushButton("OK")
        cancel_button = QPushButton("Отмена")

        button_layout.addStretch()
        button_layout.addWidget(ok_button)
        button_layout.addWidget(cancel_button)

        layout.addLayout(button_layout)
        dialog.setLayout(layout)

        def on_ok():
            try:
                values = []
                for field in fields:
                    text = field.text().strip()
                    if not text:
                        raise ValueError("Поле не может быть пустым")
                    try:
                        val = float(text.replace(',', '.'))
                        values.append(val)
                    except ValueError:
                        raise ValueError("Некорректное число")

                depth, overrun, feedrate = values
                
                print(f"Введены параметры: depth={depth}, overrun={overrun}, feedrate={feedrate}")
                
                program = self.ks_service.create_milling_program(self.ks_service.find_macro_by_type("Ноль станка"), macro, depth, overrun, feedrate)
                
                file_path, _ = QFileDialog.getSaveFileName(
                    self,
                    "Сохранить файл фрезеровки",
                    "",
                    "CNC program (*.nc)"
                )
                
                with open(file_path, "w") as file:
                    file.write(program)
                
                dialog.accept()

            except ValueError as e:
                QMessageBox.warning(dialog, "Ошибка ввода", f"Введите корректные числа\n{e}")

        ok_button.clicked.connect(on_ok)
        cancel_button.clicked.connect(dialog.reject)

        dialog.exec_()

    def show_borders_trajectory_menu(self, macro):
        """Метод для ввода параметров траектории границ"""
        dialog = QDialog(self)
        dialog.setWindowTitle("Меню параметров")
        dialog.setModal(True)
        dialog.setFixedSize(300, 180)

        layout = QVBoxLayout()

        fields = []
        labels = ["Диаметр инструмента:", "Отступ:"]
        contour_types = ["Внешний", "Внутренний"]

        validator = QDoubleValidator()
        validator.setNotation(QDoubleValidator.StandardNotation)

        for label_text in labels:
            row = QHBoxLayout()
            label = QLabel(label_text)
            line_edit = QLineEdit()
            line_edit.setPlaceholderText("0.0 мм")
            line_edit.setValidator(validator)

            row.addWidget(label)
            row.addWidget(line_edit)
            layout.addLayout(row)
            fields.append(line_edit)

        row = QHBoxLayout()
        label = QLabel("Тип контура:")
        combo = QComboBox()
        combo.addItems(contour_types)
        row.addWidget(label)
        row.addWidget(combo)
        layout.addLayout(row)
        fields.append(combo)

        button_layout = QHBoxLayout()
        ok_button = QPushButton("OK")
        cancel_button = QPushButton("Отмена")

        button_layout.addStretch()
        button_layout.addWidget(ok_button)
        button_layout.addWidget(cancel_button)

        layout.addLayout(button_layout)
        dialog.setLayout(layout)

        def on_ok():
            try:

                tool_diameter_text = fields[0].text().strip()
                offset_text = fields[1].text().strip()

                if not tool_diameter_text or not offset_text:
                    raise ValueError("Поля 'Диаметр инструмента' и 'Отступ' не могут быть пустыми")

                tool_diameter = float(tool_diameter_text.replace(',', '.'))
                offset = float(offset_text.replace(',', '.'))
                contour_type = fields[2].currentText()

                print(f"Введены параметры: tool_diameter={tool_diameter}, offset={offset}, contour_type={contour_type}")
                
                self.ks_service.create_border_trajectory(self.ks_service.find_macro_by_type("Ноль станка"), macro, tool_diameter, offset, contour_type)
                
                # program = self.ks_service.create_drilling_program(self.ks_service.find_macro_by_type("Ноль станка"), macro, depth, overrun, feedrate)
                
                # file_path, _ = QFileDialog.getSaveFileName(
                #     self,
                #     "Сохранить файл сверловки",
                #     "",
                #     "CNC program (*.nc)"
                # )
                #
                # with open(file_path, "w") as file:
                #     file.write(program)
                
                dialog.accept()

            except ValueError as e:
                QMessageBox.warning(dialog, "Ошибка ввода", f"Введите корректные числа\n{e}")

        ok_button.clicked.connect(on_ok)
        cancel_button.clicked.connect(dialog.reject)

        dialog.exec_()

    def show_tracks_trajectory_menu(self, macro):
        dialog = QDialog(self)
        dialog.setWindowTitle("Меню параметров")
        dialog.setModal(True)
        dialog.setFixedSize(300, 180)

        layout = QVBoxLayout()
        fields = []

        # Диаметр инструмента
        validator_diam = QDoubleValidator()
        validator_diam.setNotation(QDoubleValidator.StandardNotation)

        row1 = QHBoxLayout()
        label1 = QLabel("Диаметр инструмента:")
        line_edit1 = QLineEdit()
        line_edit1.setPlaceholderText("0.0 мм")
        line_edit1.setValidator(validator_diam)
        row1.addWidget(label1)
        row1.addWidget(line_edit1)
        layout.addLayout(row1)
        fields.append(line_edit1)

        # Количество линий
        validator_count = QIntValidator(1, 999)  # минимум 1, максимум 999

        row2 = QHBoxLayout()
        label2 = QLabel("Количество линий:")
        line_edit2 = QLineEdit()
        line_edit2.setPlaceholderText("1")
        line_edit2.setValidator(validator_count)
        row2.addWidget(label2)
        row2.addWidget(line_edit2)
        layout.addLayout(row2)
        fields.append(line_edit2)

        # Процент перекрытия
        validator_overlap = QDoubleValidator(0.0, 100.0, 2)
        validator_overlap.setNotation(QDoubleValidator.StandardNotation)

        row3 = QHBoxLayout()
        label3 = QLabel("Процент перекрытия:")
        line_edit3 = QLineEdit()
        line_edit3.setPlaceholderText("0.0 %")
        line_edit3.setValidator(validator_overlap)
        row3.addWidget(label3)
        row3.addWidget(line_edit3)
        layout.addLayout(row3)
        fields.append(line_edit3)

        button_layout = QHBoxLayout()
        button_layout.addStretch()
        ok_button = QPushButton("OK")
        cancel_button = QPushButton("Отмена")
        button_layout.addWidget(ok_button)
        button_layout.addWidget(cancel_button)
        layout.addLayout(button_layout)

        dialog.setLayout(layout)

        def on_ok():
            try:
                tool_diam_text = fields[0].text().strip()
                count_text = fields[1].text().strip()
                overlap_text = fields[2].text().strip()

                if not tool_diam_text or not count_text or not overlap_text:
                    QMessageBox.warning(dialog, "Ошибка", "Заполните все поля!")
                    return

                tool_diameter = float(tool_diam_text.replace(',', '.'))
                line_count = int(count_text)
                overlap_percent = float(overlap_text.replace(',', '.'))

                if tool_diameter <= 0:
                    raise ValueError("Диаметр инструмента должен быть > 0")
                if overlap_percent < 0 or overlap_percent > 100:
                    raise ValueError("Процент перекрытия должен быть от 0 до 100")

                print(f"Параметры дорожек: диаметр={tool_diameter}, линий={line_count}, перекрытие={overlap_percent}%")

                self.ks_service.create_tracks_trajectory(self.ks_service.find_macro_by_type("Ноль станка"), macro, tool_diameter, line_count, overlap_percent)

                dialog.accept()

            except ValueError as e:
                QMessageBox.warning(dialog, "Ошибка ввода", f"Некорректное значение:\n{e}")
            except Exception as e:
                QMessageBox.critical(dialog, "Ошибка", f"Не удалось обработать параметры:\n{str(e)}")

        ok_button.clicked.connect(on_ok)
        cancel_button.clicked.connect(dialog.reject)

        dialog.exec_()

    def create_start_point(self):
        self.ks_service.create_start_point()
        self.build_project_tree(self.project_name)

    def set_action_enable(self):
        self.addHolesAction.setDisabled(False)
        self.addBordersAction.setDisabled(False)
        self.addTracksAction.setDisabled(False)
        self.addMasksAction.setDisabled(False)

