import sys
from PyQt5.QtWidgets import *
from PyQt5.QtCore import Qt

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PyQt5 MainWindow Example")
        self.resize(900, 600)

        # =========================
        # Toolbar
        # =========================
        toolbar = QToolBar(self)
        self.addToolBar(toolbar)

        action_new = QAction("New", self)
        action_open = QAction("Open", self)
        action_exit = QAction("Exit", self)
        action_exit.triggered.connect(self.close)

        toolbar.addAction(action_new)
        toolbar.addAction(action_open)
        toolbar.addSeparator()
        toolbar.addAction(action_exit)

        # =========================
        # Central widget (MANDATORY)
        # =========================
        central = QWidget(self)
        self.setCentralWidget(central)

        main_layout = QVBoxLayout(central)

        # =========================
        # Top part
        # =========================
        top_widget = QWidget()
        top_layout = QHBoxLayout(top_widget)
        #top_layout.setAlignment(Qt.AlignTop)

        # --- Combo box + label ---
        combo_layout = QVBoxLayout()
        combo_label = QLabel("Options")
        combo_box = QComboBox()
        combo_box.addItems(["Option A", "Option B", "Option C"])
        combo_layout.addWidget(combo_label)
        combo_layout.addWidget(combo_box)

        # --- Radio buttons + label ---
        radio_layout = QVBoxLayout()
        radio_label = QLabel("Mode")
        radio_layout.addWidget(radio_label)

        radio_group = QButtonGroup(self)
        radio1 = QRadioButton("Mode 1")
        radio2 = QRadioButton("Mode 2")
        radio3 = QRadioButton("Mode 3")
        radio1.setChecked(True)

        radio_group.addButton(radio1)
        radio_group.addButton(radio2)
        radio_group.addButton(radio3)

        radio_layout.addWidget(radio1)
        radio_layout.addWidget(radio2)
        radio_layout.addWidget(radio3)

        # --- Line edit + label ---
        line_layout = QVBoxLayout()
        line_label = QLabel("Input")
        line_edit = QLineEdit()
        line_layout.addWidget(line_label)
        line_layout.addWidget(line_edit)

        # Add top elements
        top_layout.addLayout(combo_layout)
        top_layout.addLayout(radio_layout)
        top_layout.addLayout(line_layout)

        # =========================
        # Bottom part (table)
        # =========================
        table = QTableWidget(5, 3)
        table.setHorizontalHeaderLabels(["Col 1", "Col 2", "Col 3"])

        for r in range(5):
            for c in range(3):
                table.setItem(r, c, QTableWidgetItem(f"{r},{c}"))

        # =========================
        # Assemble main layout
        # =========================
        main_layout.addWidget(top_widget, 1)
        main_layout.addWidget(table, 3)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
