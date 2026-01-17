from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QMessageBox
)
from PyQt5.QtGui import QDoubleValidator
from domain.parameters import MillingParams

class MillingParamsDialog(QDialog):
    def __init__(self, default_params: MillingParams, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Параметры фрезеровки")
        self.setModal(True)
        self.setFixedSize(300, 180)

        layout = QVBoxLayout()

        labels = ["Глубина:", "Высота перебега:", "Скорость подачи:"]
        values = [default_params.depth, default_params.overrun, default_params.feedrate]

        self.fields = []
        validator = QDoubleValidator(0.0, 100.0, 2)
        validator.setNotation(QDoubleValidator.StandardNotation)

        for label, value in zip(labels, values):
            row = QHBoxLayout()
            row.addWidget(QLabel(label))
            line_edit = QLineEdit()
            line_edit.setText(str(value).replace('.', ','))
            line_edit.setPlaceholderText("0.0 мм")
            line_edit.setValidator(validator)
            row.addWidget(line_edit)
            layout.addLayout(row)
            self.fields.append(line_edit)

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

    def get_params(self) -> MillingParams:
        try:
            values = []
            for field in self.fields:
                text = field.text().strip().replace(',', '.')
                if not text:
                    raise ValueError("Поле не может быть пустым")
                values.append(float(text))
            return MillingParams(*values)
        except ValueError as e:
            raise ValueError(f"Некорректное число: {e}")

    @staticmethod
    def get_milling_params(defaults: MillingParams, parent=None):
        dialog = MillingParamsDialog(defaults, parent)
        if dialog.exec_() == QDialog.Accepted:
            try:
                return dialog.get_params()
            except ValueError as e:
                QMessageBox.warning(parent or dialog, "Ошибка ввода", str(e))
                return None
        return None