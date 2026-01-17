from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QMessageBox
)
from PyQt5.QtGui import QDoubleValidator, QIntValidator
from domain.parameters import TracksTrajectoryParams

class TracksTrajectoryDialog(QDialog):
    def __init__(self, default_params: TracksTrajectoryParams, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Параметры траектории дорожек")
        self.setModal(True)
        self.setFixedSize(300, 200)

        layout = QVBoxLayout()

        # Диаметр инструмента
        self.tool_diam_edit = QLineEdit()
        self.tool_diam_edit.setText(str(default_params.tool_diameter).replace('.', ','))
        self.tool_diam_edit.setValidator(QDoubleValidator(0.01, 10.0, 2))
        layout.addLayout(self._create_row("Диаметр инструмента:", self.tool_diam_edit))

        # Количество линий
        self.line_count_edit = QLineEdit()
        self.line_count_edit.setText(str(default_params.line_count))
        self.line_count_edit.setValidator(QIntValidator(1, 999))
        layout.addLayout(self._create_row("Количество линий:", self.line_count_edit))

        # Процент перекрытия
        self.overlap_edit = QLineEdit()
        self.overlap_edit.setText(str(default_params.overlap_percent).replace('.', ','))
        self.overlap_edit.setValidator(QDoubleValidator(0.0, 100.0, 2))
        layout.addLayout(self._create_row("Процент перекрытия (%):", self.overlap_edit))

        # Кнопки
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        self.ok_btn = QPushButton("OK")
        self.cancel_btn = QPushButton("Отмена")
        button_layout.addWidget(self.ok_btn)
        button_layout.addWidget(self.cancel_btn)
        layout.addLayout(button_layout)

        self.setLayout(layout)

        self.ok_btn.clicked.connect(self.accept)
        self.cancel_btn.clicked.connect(self.reject)

    def _create_row(self, label_text: str, widget):
        row = QHBoxLayout()
        row.addWidget(QLabel(label_text))
        row.addWidget(widget)
        return row

    def get_params(self) -> TracksTrajectoryParams:
        try:
            tool_diam = float(self.tool_diam_edit.text().replace(',', '.'))
            line_count = int(self.line_count_edit.text())
            overlap = float(self.overlap_edit.text().replace(',', '.'))

            if tool_diam <= 0:
                raise ValueError("Диаметр инструмента должен быть > 0")
            if not (0 <= overlap <= 100):
                raise ValueError("Процент перекрытия должен быть от 0 до 100")

            return TracksTrajectoryParams(tool_diameter=tool_diam, line_count=line_count, overlap_percent=overlap)
        except ValueError as e:
            raise ValueError(f"Некорректное значение: {e}")

    @staticmethod
    def get_track_trajectory_params(parent=None, defaults: TracksTrajectoryParams = None) -> TracksTrajectoryParams | None:
        if defaults is None:
            defaults = TracksTrajectoryParams(tool_diameter=0.2, line_count=1, overlap_percent=30.0)
        dialog = TracksTrajectoryDialog(defaults, parent)
        if dialog.exec_() == QDialog.Accepted:
            try:
                return dialog.get_params()
            except ValueError as e:
                QMessageBox.warning(parent or dialog, "Ошибка ввода", str(e))
                return None
        return None