import sys
from PyQt5.QtWidgets import QApplication
from pcb_editor import PCBEditor

if __name__ == '__main__':
    app = QApplication(sys.argv)
    editor = PCBEditor()
    editor.show()
    sys.exit(app.exec_())