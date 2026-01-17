from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QMessageBox, QComboBox
)
from PyQt5.QtGui import QDoubleValidator
from domain.parameters import BorderTrajectoryParams, ContourType

class BorderTrajectoryDialog(QDialog):
    def __init__(self, default_params: BorderTrajectoryParams, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Параметры траектории границ")
        self.setModal(True)
        self.setFixedSize(300, 180)

        layout = QVBoxLayout()

        # Диаметр инструмента
        self.tool_diam_edit = QLineEdit()
        self.tool_diam_edit.setText(str(default_params.tool_diameter).replace('.', ','))
        self.tool_diam_edit.setValidator(QDoubleValidator(0.01, 20.0, 2))
        layout.addLayout(self._create_row("Диаметр инструмента:", self.tool_diam_edit))

        # Отступ
        self.offset_edit = QLineEdit()
        self.offset_edit.setText(str(default_params.offset).replace('.', ','))
        self.offset_edit.setValidator(QDoubleValidator(0.0, 10.0, 2))
        layout.addLayout(self._create_row("Отступ:", self.offset_edit))

        # Тип контура
        self.contour_combo = QComboBox()
        self.contour_combo.addItems([ct.value for ct in ContourType])
        # Установить текущее значение
        current_index = self.contour_combo.findText(default_params.contour_type.value)
        if current_index >= 0:
            self.contour_combo.setCurrentIndex(current_index)
        layout.addLayout(self._create_row("Тип контура:", self.contour_combo))

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

    def get_params(self) -> BorderTrajectoryParams:
        try:
            tool_diam = float(self.tool_diam_edit.text().replace(',', '.'))
            offset = float(self.offset_edit.text().replace(',', '.'))
            contour_str = self.contour_combo.currentText()
            contour_type = ContourType(contour_str)

            if tool_diam <= 0:
                raise ValueError("Диаметр инструмента должен быть > 0")

            return BorderTrajectoryParams(
                tool_diameter=tool_diam,
                offset=offset,
                contour_type=contour_type
            )
        except ValueError as e:
            raise ValueError(f"Некорректное значение: {e}")

    @staticmethod
    def get_border_trajectory_params(parent=None, defaults: BorderTrajectoryParams = None) -> BorderTrajectoryParams | None:
        if defaults is None:
            defaults = BorderTrajectoryParams(
                tool_diameter=1.0,
                offset=0.1,
                contour_type=ContourType.EXTERNAL
            )
        dialog = BorderTrajectoryDialog(defaults, parent)
        if dialog.exec_() == QDialog.Accepted:
            try:
                return dialog.get_params()
            except ValueError as e:
                QMessageBox.warning(parent or dialog, "Ошибка ввода", str(e))
                return None
        return None