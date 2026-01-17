from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton, QMessageBox)
from PyQt5.QtGui import QDoubleValidator
from domain.parameters import DrillingParams

class DrillingParamsDialog(QDialog):
    def __init__(self, default_params: DrillingParams, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Параметры сверловки")
        self.setModal(True)
        self.setFixedSize(300, 180)

        layout = QVBoxLayout()
        self.fields = []

        labels = ["Глубина:", "Высота перебега:", "Скорость подачи:"]
        values = [default_params.depth, default_params.overrun, default_params.feedrate]

        validator = QDoubleValidator()
        validator.setNotation(QDoubleValidator.StandardNotation)

        for label_text, value in zip(labels, values):
            row = QHBoxLayout()
            label = QLabel(label_text)
            line_edit = QLineEdit()
            line_edit.setText(str(value).replace('.', ','))
            line_edit.setPlaceholderText("0.0 мм")
            line_edit.setValidator(validator)
            row.addWidget(label)
            row.addWidget(line_edit)
            layout.addLayout(row)
            self.fields.append(line_edit)

        button_layout = QHBoxLayout()
        button_layout.addStretch()
        self.ok_button = QPushButton("OK")
        self.cancel_button = QPushButton("Отмена")
        button_layout.addWidget(self.ok_button)
        button_layout.addWidget(self.cancel_button)
        layout.addLayout(button_layout)

        self.setLayout(layout)

        self.ok_button.clicked.connect(self.accept)
        self.cancel_button.clicked.connect(self.reject)

    def get_params(self) -> DrillingParams:
        try:
            values = []
            for field in self.fields:
                text = field.text().strip().replace(',', '.')
                if not text:
                    raise ValueError("Поле не может быть пустым")
                values.append(float(text))
            return DrillingParams(*values)
        except ValueError as e:
            raise ValueError(f"Некорректное число: {e}")

    @staticmethod
    def get_drilling_params(defaults: DrillingParams, parent=None):
        dialog = DrillingParamsDialog(defaults, parent)
        if dialog.exec_() == QDialog.Accepted:
            try:
                return dialog.get_params()
            except ValueError as e:
                QMessageBox.warning(parent or dialog, "Ошибка ввода", str(e))
                return None
        return None