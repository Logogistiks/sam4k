import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QAction, QTextEdit
from PyQt5.QtGui import QIcon

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("PyQt5 MainWindow")
        self.setGeometry(100, 100, 800, 600)

        # Central widget
        self.setCentralWidget(QTextEdit())

        # Actions
        open_action = QAction(QIcon("testing/Ringspruch.jpg"), "Open", self)
        open_action.setShortcut("Ctrl+O")
        open_action.triggered.connect(self.open_file)

        exit_action = QAction(QIcon("icons/exit.png"), "Exit", self)
        exit_action.setShortcut("Ctrl+Q")
        exit_action.triggered.connect(self.close)

        # Toolbar
        toolbar = self.addToolBar("Main")
        toolbar.addAction(open_action)
        toolbar.addAction(exit_action)

    def open_file(self):
        print("Open action triggered")

app = QApplication(sys.argv)
window = MainWindow()
window.show()
sys.exit(app.exec_())
