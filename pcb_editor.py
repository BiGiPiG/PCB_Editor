import os
from pathlib import Path

from PyQt5.QtWidgets import (QMainWindow, QAction, QTextEdit, QStatusBar, QMessageBox, QFileDialog,
                             QDialog, QVBoxLayout, QLabel, QLineEdit, QHBoxLayout, QPushButton)
from kompas_service import KompasService


class PCBEditor(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.ks_service = KompasService()

    def initUI(self):
        self.setWindowTitle('Редактор печатных плат')
        self.setGeometry(100, 100, 1000, 700)

        self.canvas = QTextEdit()
        self.setCentralWidget(self.canvas)

        self.create_menus()

        self.statusBar = QStatusBar()
        self.setStatusBar(self.statusBar)
        self.statusBar.showMessage("Готово")

    def create_menus(self):
        # Меню Файл
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

        addTracksAction = QAction('Добавить дорожки', self)
        addTracksAction.triggered.connect(self.add_tracks)
        fileMenu.addAction(addTracksAction)

        addBordersAction = QAction('Добавить границы', self)
        addBordersAction.triggered.connect(self.add_borders)
        fileMenu.addAction(addBordersAction)

        addHolesAction = QAction('Добавить отверстия', self)
        addHolesAction.triggered.connect(self.add_holes)
        fileMenu.addAction(addHolesAction)

        exitAction = QAction('Выход', self)
        exitAction.setShortcut('Ctrl+Q')
        exitAction.triggered.connect(self.close)
        fileMenu.addAction(exitAction)

    def open_project(self):
        print("open project")
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Открыть проект",
            "",
            "PCB Project Files (*.pcbprj)"
        )
        if file_path:
            self.ks_service.open_fragment(file_path.replace(".pcbprj", ".frw"))

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
            project_name = name_input.text().strip()
            project_folder = folder_input.text().strip()

            if not project_name:
                QMessageBox.warning(dlg, "Ошибка", "Введите название проекта!")
                return

            if not project_folder:
                QMessageBox.warning(dlg, "Ошибка", "Выберите папку для проекта!")
                return

            project_path = Path(project_folder) / f"{project_name}"

            if not os.path.exists(project_path):
                QMessageBox.warning(dlg, "Ошибка", f"Папка '{project_path}' не существует!")
                return

            try:
                os.makedirs(project_path, exist_ok=True)
                frw_path = project_path / f"{project_name}.frw"
                pcbprj_path = project_path / f"{project_name}.pcbprj"
                print(frw_path)
                if os.path.exists(frw_path):
                    response = QMessageBox.question(
                        dlg,
                        "Подтверждение",
                        "Проект уже существует. Перезаписать?",
                        QMessageBox.Yes | QMessageBox.No,
                        QMessageBox.No
                    )
                    if response != QMessageBox.Yes:
                        return

                self.ks_service.create_fragment(frw_path)
                with open(pcbprj_path, 'w'): pass

                self.statusBar.showMessage(f"Создан проект: {project_path}")
                QMessageBox.information(dlg, "Успех", f"Проект '{project_name}' создан успешно!")
                dlg.accept()

            except Exception as e:
                QMessageBox.critical(dlg, "Ошибка", f"Не удалось создать проект: {str(e)}")

        def cancel():
            dlg.reject()

        folder_button.clicked.connect(browse_folder)
        create_button.clicked.connect(create_project)
        cancel_button.clicked.connect(cancel)

        dlg.exec()

    def add_tracks(self):
        self.statusBar.showMessage("Режим добавления дорожек")
        # TODO функционал добавления дорожек

    def add_borders(self):
        self.statusBar.showMessage("Режим добавления границ")
        # TODO функционал добавления границ

    def add_holes(self):
        self.statusBar.showMessage("Режим добавления отверстий")
        # TODO функционал добавления отверстий
