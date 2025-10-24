import os
import traceback

from PyQt5.QtGui import QDoubleValidator

from drillingFileReader import drillingFileReader
from dxfFileReader import DXFReader
from pathlib import Path

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (QMainWindow, QAction, QStatusBar, QMessageBox, QFileDialog,
                             QDialog, QVBoxLayout, QLabel, QLineEdit, QHBoxLayout, QPushButton, QTreeWidget,
                             QTreeWidgetItem, QMenu)
from kompas_service import KompasService


class PCBEditor(QMainWindow):
    def __init__(self):
        super().__init__()
        self.addTracksAction = None
        self.project_name = None
        self.tree_view = None
        self.ks_service = KompasService()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('Редактор печатных плат')
        self.setGeometry(100, 100, 500, 300)

        self.tree_view = QTreeWidget()
        self.tree_view.setHeaderLabel("")
        self.tree_view.setContextMenuPolicy(Qt.CustomContextMenu)
        self.tree_view.customContextMenuRequested.connect(self.show_context_menu)
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
        delete_action.triggered.connect(lambda: self.ks_service.delete_macro(item.text(0)))

        context_menu.addAction(rename_action)
        context_menu.addAction(delete_action)
        if item.text(0) == "Отверстия":
            control_program_action.triggered.connect(lambda: self.show_drill_menu(item.data(1, 0)))
            context_menu.addAction(control_program_action)
        if item.text(0) == "Границы":
            trajectory_action.triggered.connect(lambda: self.show_drill_menu(item.data(1, 0)))
            context_menu.addAction(trajectory_action)
        if item.text(0) == "Дорожки":
            trajectory_action.triggered.connect(lambda: self.show_drill_menu(item.data(1, 0)))
            context_menu.addAction(trajectory_action)

        context_menu.exec_(self.tree_view.viewport().mapToGlobal(position))


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
            "Tracks Files (*.drl)"
        )

        if not file_path:
            self.statusBar.showMessage("Файл не выбран")
            return

        try:
            border_data = DXFReader.readFile(file_path)
            self.ks_service.draw_border(border_data)
            self.statusBar.showMessage(f"Дорожки загружены из: {os.path.basename(file_path)}")
        except Exception as e:
            self.statusBar.showMessage(f"Ошибка загрузки: {str(e)}")
            print(f"Ошибка при чтении файла дорожек: {e}")

    def add_borders(self):
        self.statusBar.showMessage("Режим добавления границ")
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Открыть файл границ",
            "",
            "Borders Files (*.drl)"
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

    def create_start_point(self):
        self.ks_service.create_start_point()
        self.build_project_tree(self.project_name)

    def set_action_enable(self):
        self.addHolesAction.setDisabled(False)
        self.addBordersAction.setDisabled(False)
        self.addTracksAction.setDisabled(False)

