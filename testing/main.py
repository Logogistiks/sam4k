import sys
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *

import os
os.chdir("testing")

class AutoResizeLineEdit(QLineEdit):
    def __init__(self, min_width=150, padding=10):
        super().__init__()
        self._min_width = min_width
        self._padding = padding
        self.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self.textChanged.connect(self._update_width)
        self._update_width()

    def _update_width(self):
        font_metrics = QFontMetrics(self.font())
        text_width = font_metrics.horizontalAdvance(self.text() or " ")
        new_width = max(self._min_width, text_width + self._padding)
        self.setMinimumWidth(new_width)

class LabelComboBox(QWidget):
    def __init__(self, label_text, items, default_index=0, editable=False):
        super().__init__()
        self.LClayout = QVBoxLayout(self)
        self.LClayout.setAlignment(Qt.AlignTop)

        self.LClabel = QLabel(label_text)
        self.LCcombo_box = QComboBox()
        self.LCcombo_box.addItems(items)
        self.LCcombo_box.setCurrentIndex(default_index)
        self.LCcombo_box.setEditable(editable)

        self.LClayout.addWidget(self.LClabel)
        self.LClayout.addWidget(self.LCcombo_box)

class LabelRadioGroup(QWidget):
    def __init__(self, label_text, options, default_index=0):
        super().__init__()
        self.LRlayout = QVBoxLayout(self)
        self.LRlayout.setAlignment(Qt.AlignTop)

        self.LRlabel = QLabel(label_text)
        self.LRlayout.addWidget(self.LRlabel)

        self.LRbutton_group = QButtonGroup(self)
        for i, option in enumerate(options):
            radio_button = QRadioButton(option)
            if i == default_index:
                radio_button.setChecked(True)
            self.LRbutton_group.addButton(radio_button)
            self.LRlayout.addWidget(radio_button)

class LabelLineEdit(QWidget):
    def __init__(self, label_text, min_width=150, grow=False):
        super().__init__()
        self.LLlayout = QVBoxLayout(self)
        self.LLlayout.setAlignment(Qt.AlignTop)

        self.LLlabel = QLabel(label_text)
        if grow:
            self.LLline_edit = AutoResizeLineEdit(min_width=min_width)
        else:
            self.LLline_edit = QLineEdit()
            self.LLline_edit.setMinimumWidth(min_width)

        self.LLlayout.addWidget(self.LLlabel)
        self.LLlayout.addWidget(self.LLline_edit)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("SAM4k")
        self.setWindowIcon(QIcon("icons/sam4k.svg"))
        self.resize(900, 600)

        # =========================
        # region: Toolbar
        # =========================

        toolbar = QToolBar(self)
        toolbar.setMovable(False)
        toolbar.setFloatable(False)
        toolbar.setContextMenuPolicy(Qt.PreventContextMenu)
        self.addToolBar(toolbar)

        action_run = QAction(QIcon("icons/run.svg"), "Run", self)
        action_run.setShortcut("Ctrl+R")
        action_run.triggered.connect(lambda: print("Run action triggered"))
        toolbar.addAction(action_run)

        action_next = QAction(QIcon("icons/next.svg"), "Next", self)
        action_next.setShortcut("Ctrl+N")
        action_next.triggered.connect(lambda: print("Next action triggered"))
        toolbar.addAction(action_next)

        action_save = QAction(QIcon("icons/save.svg"), "Save", self)
        action_save.setShortcut("Ctrl+S")
        action_save.triggered.connect(lambda: print("Save action triggered"))
        toolbar.addAction(action_save)

        action_clear = QAction(QIcon("icons/clear.svg"), "Clear", self)
        action_clear.setShortcut("Ctrl+q")
        action_clear.triggered.connect(lambda: print("Clear action triggered"))
        toolbar.addAction(action_clear)

        action_settings = QAction(QIcon("icons/settings.svg"), "Settings", self)
        action_settings.triggered.connect(lambda: print("Settings action triggered"))
        toolbar.addAction(action_settings)

        # endregion: Toolbar
        # =========================

        # =========================
        # region: CentralWidget
        # =========================

        central = QWidget(self)
        self.setCentralWidget(central)

        layout_main = QVBoxLayout(central)

        # endregion: CentralWidget
        # =========================

        # =========================
        # region: Top Part
        # =========================

        top = QGroupBox("Vorkonfiguration")
        top.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed) # no vertical strech
        layout_top = QHBoxLayout(top)
        layout_top.setAlignment(Qt.AlignLeft)
        layout_top.setSizeConstraint(QLayout.SetMinimumSize) # fit content
        layout_top.setSpacing(50)

        shots_per_strip = LabelComboBox("Schüsse pro Streifen", ["10", "5", "2", "1"])
        shots_per_strip.LCcombo_box.currentIndexChanged.connect(self.index_changed)
        layout_top.addWidget(shots_per_strip)

        savemode = LabelRadioGroup("Speichermodus", ["mit Teiler", "ohne Teiler", "einzeln mit Teiler, Gesamt ohne"], default_index=0)
        layout_top.addWidget(savemode)

        name = LabelLineEdit("Name des Schützen", grow=True)
        layout_top.addWidget(name)

        layout_main.addWidget(top)

        # endregion: Top Part
        # =========================

        # =========================
        # region: Bottom Part
        # =========================

        bottom = QGroupBox("Ergebnisse")
        layout_bottom = QVBoxLayout(bottom)

        result_table = QTableWidget()
        layout_bottom.addWidget(result_table)

        layout_main.addWidget(bottom)

        # endregion: Bottom Part
        # =========================

        # =========================
        # region: Statusbar
        # =========================

        statusbar = self.statusBar()
        statusbar.showMessage("Bereit", 5000)

        # endregion: Statusbar
        # =========================

    def index_changed(self, i):
        print(i)

    def showEvent(self, event):
        print("Window is shown")
        super().showEvent(event)

    def closeEvent(self, event):
        print("Window is closing")
        event.accept()

app = QApplication(sys.argv)

window = MainWindow()
window.show()

app.exec()

# https://www.pythonguis.com/pyqt5-tutorial/
# https://www.pythonguis.com/tutorials/pyqt-layouts/